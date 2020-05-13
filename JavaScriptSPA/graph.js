// Helper function to call MS Graph API endpoint 
// using authorization bearer token scheme
//d

function callMSGraph(endpoint, accessToken, callback) {
  const headers = new Headers();
  const bearer = `Bearer ${accessToken}`;
  console.log(bearer)

  window.localStorage.setItem('token', bearer);
 
 
  headers.append("Authorization", bearer);

  const options = {
      method: "GET",
      headers: headers
    };

  console.log('request made to Graph API at: ' + new Date().toString());
  
  fetch(endpoint, options)
    .then(response => response.json())
    .then(response => callback(response, endpoint,accessToken))
    .catch(error => console.log(error))
}

function seeProfile() {
  if (myMSALObj.getAccount()) {
    getTokenPopup(loginRequest)
      .then(response => {
        callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
        profileButton.style.display = 'none';
      }).catch(error => {
        console.log(error);
      });
  }
}

function readMail() {
  if (myMSALObj.getAccount()) {
      getTokenPopup(tokenRequest)
        .then(response => {
          callMSGraph(graphConfig.graphbookingBusinessesAppointments, response.accessToken, updateUI);
        //  alert(response.accessToken)
          mailButton.style.display = 'none';
        }).catch(error => {
          console.log(error);
        });
  }
}

function readBookings() {
   
  if (myMSALObj.getAccount()) {
      getTokenPopup(tokenRequest)
        .then(response => {
      //   alert("zz"+response.accessToken)
      //    alert(graphConfig.graphbookingBusinessesEndpoint)
       callMSGraph('https://graph.microsoft.com/beta//bookingBusinesses', response.accessToken, updateUI);
        }).catch(error => {
          console.log(error);
        });
  }
  
}