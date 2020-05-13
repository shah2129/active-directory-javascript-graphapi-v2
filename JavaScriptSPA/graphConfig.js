  // Add here the endpoints for MS Graph API services you would like to use.
  const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages",
    graphbookingBusinessesEndpoint: "https://graph.microsoft.com/v1.0/me/bookingBusinesses",
    graphbookingBusinessesAppointments: "https://graph.microsoft.com/beta/bookingBusinesses"
  };

  // Add here scopes for access token to be used at MS Graph API endpoints.
  const tokenRequest = {
      scopes: ["Bookings.Read.All","Calendars.ReadWrite.Shared"]
  };