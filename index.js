// module.exports = async function (context, req) {
//     context.log('JavaScript HTTP trigger function processed a request.');

//     const name = (req.query.name || (req.body && req.body.name));
//     const responseMessage = name
//         ? "Hello, " + name + ". This HTTP triggered function executed successfully."
//         : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

//     context.res = {
//         // status: 200, /* Defaults to 200 */
//         body: responseMessage
//     };
// }


var request = require('request');
var msal = require('@azure/msal-node');



function sendMail(token, email_address, mailbox_message_id) {
  return new Promise((resolve, reject) => {
    const options = {
      method: 'POST',
      url: 'https://graph.microsoft.com/v1.0/users/' + email_address + '/' + mailbox_message_id + '/send',
      headers: {
        'Authorization': 'Bearer ' + token,
        'content-type': 'application/json'
      }
    };
    
    request(options, (error, response, body) => {
      const result = JSON.parse(body);
      if (!error && response.statusCode == 204) {
        resolve(result.value);
      } else {
        reject(result);
      }
    });
  });
}

function listMail(token, email_address) {
    return new Promise((resolve, reject) => {
      const options = {
        method: 'GET',
        url: 'https://graph.microsoft.com/v1.0/users/' + email_address + '/messages',
        headers: {
          'Authorization': 'Bearer ' + token,
          'content-type': 'application/json'
        }
      };
      
      request(options, (error, response, body) => {
        const result = JSON.parse(body);
        if (!error ) {
          resolve(result.value, response.statusCode);
        } else {
          reject(result);
        }
      });
    });
  }

function getToken(){
  return new Promise((resolve, reject) => {
    const msalConfig = {
      auth: {
          clientId: process.env["CLIENT_ID"],
          authority: `https://login.microsoftonline.com/${process.env["TENANT"]}`,
          clientSecret: process.env["CLIENT_SECRET"],
      }
    };
    // Create msal application object
    const cca = new msal.ConfidentialClientApplication(msalConfig);
    // With client credentials flows permissions need to be granted in the portal by a tenant administrator.
    // The scope is always in the format "<resource>/.default"
    const clientCredentialRequest = {
        scopes: ["https://graph.microsoft.com/.default"],
    };
    cca.acquireTokenByClientCredential(clientCredentialRequest).then((response) => {
        resolve(response.accessToken);
    }).catch((error) => {
        reject(error);
    });
  })
}

module.exports = function (context, req) {
    context.log('Starting function');
    if (req.query.email_address) {
    const email_address = req.query.email_address;
      getToken().then(token => {
        listMail(token, email_address)
          .then((result, statusCode) => {
            context.res = {
              status: statusCode,
              body: JSON.stringify(result),
              headers: {
                'Content-Type': 'application/json'
              }
            };
            context.done();
          }).catch(() => {
            context.log('An error occurred while asking MS Graph API');
            context.done();
          });
      }).catch(()=>{
        context.res = {
            status: 400,
            body: "Impossible to get Token"
          };
          context.done();  
      });
    } else {
      context.res = {
        status: 400,
        body: "Please pass an email_address on the query string"
      };
      context.done();        
    }
  };