import React, { Component } from "react";
import "./App.css";
import { Text } from "office-ui-fabric-react";
import { GraphFileBrowser } from "@microsoft/file-browser";
import { UserAgentApplication } from "msal";
import range from "lodash/range";

const CLIENT_ID = "a9aa1338-3439-4ac7-b6bc-a0bc1c4d1c9b";
const SCOPES = ["Files.ReadWrite.All", "user.read"];

class App extends Component {
  constructor(props) {
    super(props);

    this.state = {
      token: null
    };

    this.msal = new UserAgentApplication(CLIENT_ID);
  }

  componentDidMount() {
    this._tryGetMsalToken();
  }

  render() {
    const { token } = this.state;

    return (
      <div className='App'>
        <button onClick={() => this.execute()}>execute</button>
        <button onClick={() => this.removePermissions()}>
          remove permissions
        </button>
        <Text variant='xxLarge'>File Browser Demo</Text>
        {token ? (
          <GraphFileBrowser
            getAuthenticationToken={() => Promise.resolve(token)}
            onSuccess={alert}
            onRenderCancelButton={() => null}
            onRenderSuccessButton={() => null}
          />
        ) : (
          <Text block onClick={this._getMsalToken}>
            [log in]
          </Text>
        )}
      </div>
    );
  }
  async http(url, opts) {
    return new Promise((res, rej) => {
      fetch(url, opts)
        .then(response => {
          response
            .text()
            .then(json => {
              console.log("json", json);
              json ? res(JSON.parse(json)) : res(json);
            })
            .catch(e => rej(e));
        })
        .catch(e => rej(e));
    });
  }
  get(url) {
    const { token } = this.state;
    const headers = new Headers();
    headers.append("Content-Type", "application/json");
    headers.append("Authorization", `bearer ${token}`);
    return this.http(url, {
      method: "GET",
      headers
    });
  }
  post(url, body) {
    const { token } = this.state;
    const headers = new Headers();
    headers.append("Content-Type", "application/json");
    headers.append("Authorization", `bearer ${token}`);
    return this.http(url, {
      method: "POST",
      headers,
      body
    });
  }
  async execute() {
    const folderName = "proveeee2";
    const projectFolderID = "01ZLZ7JBGBYHC2BHFRIZF2DQ7NUSSEFKRR";
    const templateID = "01ZLZ7JBHVKIUB2JRLERF3VAFPF4IPIMOV";

    const responseFolders = await this.get(
      `https://graph.microsoft.com/v1.0/me/drive/items/${projectFolderID}/children`
    );
    console.log("responseFolders", responseFolders);
    const actualFolder = responseFolders.value.find(
      folder => folder.name.indexOf(folderName) !== -1
    );
    let actualFolderID;
    if (actualFolder) {
      actualFolderID = actualFolder.id;
    } else {
      const newActualFolderResponse = await this.post(
        `https://graph.microsoft.com/v1.0/me/drive/items/${projectFolderID}/children`,
        JSON.stringify({
          "@microsoft.graph.conflictBehavior": "rename",
          folder: {},
          name: folderName
        })
      );
      actualFolderID = newActualFolderResponse.id;
    }
    console.log("actualFolderID", actualFolderID);
    await Promise.all(
      range(6).map(i => {
        return this.post(
          `https://graph.microsoft.com/v1.0/me/drive/items/${templateID}/copy`,
          JSON.stringify({
            parentReference: {
              id: actualFolderID
            },
            name: `${folderName}_${i}.xlsx`
          })
        );
      })
    );

    const filesRaw = await this.get(
      `https://graph.microsoft.com/v1.0/me/drive/items/${actualFolderID}/children`
    );
    const files = [];
    for (const fileRaw of filesRaw.value) {
      const shareUrlRaw = await this.post(
        `https://graph.microsoft.com/v1.0/me/drive/items/${fileRaw.id}/createLink`,
        JSON.stringify({
          type: "edit",
          scope: "anonymous"
        })
      );
      files.push({
        id: fileRaw.id,
        sharePermissionId: shareUrlRaw.id,
        shareUrl: shareUrlRaw.link.webUrl,
        url: fileRaw["@microsoft.graph.downloadUrl"],
        name: fileRaw.name
      });
    }
  }

  _acquireAccessToken = () => {
    this.msal.acquireTokenSilent(SCOPES).then(
      token => {
        this.setState({ token });
      },
      err => {
        this.msal.acquireTokenRedirect(SCOPES);
      }
    );
  };

  _loginPromptAndAuthenticate = () => {
    this.msal.loginPopup(SCOPES).then(idToken => {
      this._acquireAccessToken();
    });
  };

  _tryGetMsalToken = () => {
    const user = this.msal.getUser();

    if (user) {
      this._acquireAccessToken();
    }
  };

  _getMsalToken = () => {
    const user = this.msal.getUser();

    if (user) {
      this._acquireAccessToken();
    } else {
      this._loginPromptAndAuthenticate();
    }
  };
}

export default App;
