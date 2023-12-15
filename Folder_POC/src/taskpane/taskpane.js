import { PublicClientApplication } from "@azure/msal-browser";
(function () {
  "use strict";

  let msalInstance;
  let accessTOKEN;
  let userID;
  let selectedClient;
  let selectedMatter;


  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      const container = document.getElementById("container");
      const htmlComponent = `
      <div>
          <div id="authContainer">
            <button id="loginButton">Sign In</button>
          </div>
          <div id="folderContainer">
            <p id="userInfo"></p>
            <button id="listarFolderButton">Listar Folder</button>
            <label for="clientDropdown">Cliente:</label>
            <select id="clientDropdown">
              <option value="">Seleccione un cliente</option>
              <option value="clientA">Cliente A</option>
              <option value="clientB">Cliente B</option>
              <option value="clientC">Cliente C</option>
            </select>

            <label for="matterDropdown">Matter:</label>
            <select id="matterDropdown">
              <option value="">Seleccione un matter</option>
            </select>

            <button id="createFolderButton">Crear Carpeta</button>

            <p id="errorMessage" style="color: red; margin-top: 10px;"></p>
          </div>
        </div>
      `;

      // Renderiza el componente HTML en el elemento
       // Renderiza el componente HTML en el elemento
       container.innerHTML = htmlComponent;

       const clientDropdown = document.getElementById("clientDropdown");
       const matterDropdown = document.getElementById("matterDropdown");
   
       document.getElementById("loginButton").onclick = login;
       document.getElementById("listarFolderButton").onclick = getlistMialBox;
       document.getElementById("createFolderButton").onclick = createFolder;
       clientDropdown.addEventListener("change", loadMatters);
   
       // Agrega un evento de cambio al dropdown de matters
       matterDropdown.addEventListener("change", function () {
         selectedMatter = matterDropdown.value;
       });
    }
  });

  async function loadMatters() {
    selectedClient = document.getElementById("clientDropdown").value;

    const matters = getMattersForClient(selectedClient);

    const matterDropdown = document.getElementById("matterDropdown");
    matterDropdown.innerHTML = '<option value="">Seleccione un matter</option>';

    matters.forEach((matter) => {
      const option = document.createElement("option");
      option.value = matter;
      option.text = matter;
      matterDropdown.add(option);
    });
    selectedMatter = matterDropdown.value;
  }

  async function login() {
    try {
      msalInstance = new PublicClientApplication({
        auth: {
          clientId: "ClientID",
          authority: "https://login.microsoftonline.com/TENANT-ID",
          redirectUri: "https://localhost:3000",
          scopes: ["openid", "profile", "email", "https://graph.microsoft.com/mail.readwrite"],
        },
      });

      const authResponse = await msalInstance.loginPopup();
      updateUI(authResponse.account, authResponse.idTokenClaims.oid);
      accessTOKEN = authResponse.accessToken;
      userID = authResponse.idTokenClaims.oid;
      const account = msalInstance.getAllAccounts()[0];
      console.log("JPI account ", account )
    } catch (error) {
      console.error("Error en la autenticación:", error);
    }
  }



  async function getlistMialBox() {
   
    var apiUrl = `https://graph.microsoft.com/v1.0/users/${userID}/mailFolders`;
    try {
      // Realizar la llamada a Microsoft Graph API utilizando fetch y async/await
      const response = await fetch(apiUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessTOKEN}`,
          "Content-Type": "application/json",
        },
      });

      if (response.ok) {
        console.log("Data", response);
        clearError();
      } else {
        const error = await response.json();
        console.error("Error:", error);
        displayError(`Error: ${error.error.message}`);
      }
    } catch (error) {
      const error2 = await response.json();
      console.log("error", error2);
      displayError(`Error: ${error2.error.message}`);
    }
  }

  function displayError(errorMessage) {
    const errorMessageElement = document.getElementById("errorMessage");
    errorMessageElement.textContent = errorMessage;
  }

  function clearError() {
    const errorMessageElement = document.getElementById("errorMessage");
    errorMessageElement.textContent = "";
  }

  async function createFolder() {
    try {
     if (!selectedClient || !selectedMatter) {
      displayError("Seleccione un cliente y un matter antes de crear la carpeta.");
      return;
    }
    const folderName = `${selectedClient}-${selectedMatter}`;

      const accessToken = await getAccessToken();

      const request = {
        displayName: folderName,
        isHidden: false,
      };
      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userID}/mailFolders`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(request),
      });

      if (response.ok) {
        console.log("Creation successful.");
        clearError();
      } else {
        const error = await response.json();
        console.error("Error:", error);
        displayError(`Error: ${error.error.message}`);
      }
    } catch (error) {
      console.error("Error:", error);
      displayError(`Unexpected error: ${error.error.message}`);
    }
  }

  function getMattersForClient(client) {
    switch (client) {
      case "clientA":
        return ["A1", "A2", "A3"];
      case "clientB":
        return ["B1", "B2", "B3"];
      case "clientC":
        return ["C1", "C2", "C3"];
      default:
        return [];
    }
  }

  function updateUI(account, userid) {
    const authContainer = document.getElementById("authContainer");
    const folderContainer = document.getElementById("folderContainer");
    const userInfo = document.getElementById("userInfo");

    authContainer.style.display = "none";
    folderContainer.style.display = "block";

    userInfo.innerHTML = `Usuario: ${account.name}, User ID: ${userid}`;
  }

  async function getAccessToken() {
    if (accessTOKEN) {
      return accessTOKEN;
    } else {
      const msalConfig = {
        // auth: {
        //   clientId: '60bac751-30b3-4653-a688-c6d5fbcdf077', // Reemplaza con tu ID de cliente de Azure AD
        //   authority: 'https://login.microsoftonline.com/17699375-6cb8-4311-a6ed-63b7d11a56b4', // Reemplaza con la URL de autoridad de tu inquilino de Azure AD
        //   redirectUri: 'https://localhost:3000', // Reemplaza con la URL de redirección de tu aplicación,
        //   scopes: ["openid", "profile", "email", "https://graph.microsoft.com/mail.readwrite"], // Agrega lo // Agrega los scopes necesarios
        // },
        auth: {
          clientId: "ClientID", // Reemplaza con tu ID de cliente de Azure AD
          authority: "https://login.microsoftonline.com/TENANT-ID", // Reemplaza con la URL de autoridad de tu inquilino de Azure AD
          redirectUri: "https://localhost:3000", // Reemplaza con la URL de redirección de tu aplicación,
          scopes: ["openid", "profile", "email", "https://graph.microsoft.com/mail.readwrite"], // Agrega lo // Agrega los scopes necesarios
        },
      };

      const msalInstance = new PublicClientApplication(msalConfig);

      // Usa MSAL para obtener el token de acceso
      const authResponse = await msalInstance.loginPopup();
      console.log("JPI ", authResponse);

      // Maneja la respuesta de autenticación y devuelve el token de acceso
      return authResponse.accessToken;
    }
  }
})();
