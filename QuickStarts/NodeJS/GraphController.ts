import fetch  from 'node-fetch';

const baseUrl = "https://graph.microsoft.com/beta/sites";

const HTTP_CODE = {
  SUCCESS: 200,
  CREATED: 201,
  NO_CONTENT: 204,
}

export interface IPage {
  id?: string;
  name: string;
  title?: string;
  description?: string;
  lastModifiedDateTime?: string;
  pageLayout?: 'article' | 'microsoftReserved' | 'home';
  thumbnailWebUrl?: string;
  promotionKind?: 'page' | 'newsPost' | 'microsoftReserved';
  showComments?: boolean;
  showRecommendedPages?: boolean;
}

/**
 * Example controller to handle Microsoft Graph Pages API requests.
 *
 * @export
 * @class GraphPagesAPI
 */
export default class GraphPagesAPI {
  appId: string;
  appSecret: string;
  tenantId: string;
  token: {
    access_token?: string;
  };

  constructor(config) {
    this.appId = config.appId;
    this.appSecret = config.appSecret;
    this.tenantId = config.tenantId;
    this.token = {};
  }

  storeToken = (token: { access_token: string }) => {
    this.token = token;
  }

  setToken = (token: string) => {
    this.token.access_token = token;
  }

  /**
   * Creates a POST request to receive OAuth authorization token for your tenant ID.
   *
   * @memberof GraphPagesAPI
   */
  getAuthenticationToken = async (): Promise<{ access_token: string }> => {
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const options = {
      method: 'POST',
      form: {
        client_id: this.appId,
        client_secret: this.appSecret,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default'
      }
    };
    const res = await fetch(url, options);
    const resBody: { access_token: string } = await res.json();
    return resBody;
  }

  /**
   * Creates a GET request to retrieve all pages on target site.
   *
   * @param {string} siteId
   * @memberof GraphPagesAPI
   */
  listPages = async (siteId: string): Promise<{ value: IPage[] }> => {
    const url = `${baseUrl}/${siteId}/pages`;
    const options = {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${this.token.access_token}`,
      },
    };
    const res = await fetch(url, options);
    if (res.status !== HTTP_CODE.SUCCESS) {
      throw new Error(`Error: ${res.status} ${res.statusText}, ${await res.text()}`);
    }
    const resBody: { value: IPage[] } = await res.json();
    return resBody;
  }

  /**
   * Creates a GET request to get a specific page from target site.
   * 
   * @param {string} siteId
   * @param {string} pageId
   * @memberof GraphPagesAPI
   */
  getPage = async (siteId: string, pageId: string): Promise<IPage> => {
    const url = `${baseUrl}/${siteId}/pages/${pageId}?expand=canvasLayout`;
    const options = {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${this.token.access_token}`,
      },
    };
    const res = await fetch(url, options);
    if (res.status !== HTTP_CODE.SUCCESS) {
      throw new Error(`Error: ${res.status} ${res.statusText}, ${await res.text()}`);
    }
    const resBody: IPage = await res.json();
    return resBody;
  }

  /**
   * Creates a POST request to create a new page with payload in the target site.
   *
   * @param {string} siteId
   * @param {IPage} pagePayload
   * @memberof GraphPagesAPI
   */
  createPage = async (siteId: string, pagePayload: IPage): Promise<IPage> => {
    const url = `${baseUrl}/${siteId}/pages`;
    const options = {
      method: 'POST',
      body: JSON.stringify(pagePayload),
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${this.token.access_token}`,
      },
    };
    const res = await fetch(url, options);
    if (res.status !== HTTP_CODE.CREATED) {
      throw new Error(`Error: ${res.status} ${res.statusText}, ${await res.text()}`);
    }
    const resBody: IPage = await res.json();
    return resBody;
  }

  /**
   * Creates a POST request to publish a specific page in the target site.
   *
   * @param {string} siteId
   * @param {string} pageId
   * @memberof GraphPagesAPI
   */
  publishPage = async (siteId: string, pageId: string): Promise<void> => {
    const url = `${baseUrl}/${siteId}/pages/${pageId}/publish`;
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${this.token.access_token}`,
      },
    };
    const res = await fetch(url, options);
    if (res.status !== HTTP_CODE.NO_CONTENT) {
      throw new Error(`Error: ${res.status} ${res.statusText}, ${await res.text()}`);
    }
  }

  /**
   * Creates a DELETE request to delete a specific page in the target site.
   *
   * @param {string} siteId
   * @param {string} pageId
   * @memberof GraphPagesAPI
   */
  deletePage = async (siteId: string, pageId: string): Promise<void> => {
    const url = `${baseUrl}/${siteId}/pages/${pageId}`;
    const options = {
      method: 'DELETE',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${this.token.access_token}`,
      },
    };
    const res = await fetch(url, options);
    if (res.status !== HTTP_CODE.NO_CONTENT) {
      throw new Error(`Error: ${res.status} ${res.statusText}, ${await res.text()}`);
    }
  }

  /**
   * Creates a POST request to update a specific page with payload in the target site.
   *
   * @param {string} siteId
   * @param {string} pageId
   * @param {Partial<IPage>} pagePayload
   * @memberof GraphPagesAPI
   */
  updatePage = async (siteId: string, pageId: string, pagePayload: Partial<IPage>): Promise<IPage> => {
    const url = `${baseUrl}/${siteId}/pages/${pageId}`;
    const options = {
      method: 'PATCH',
      body: JSON.stringify(pagePayload),
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${this.token.access_token}`,
      },
    };
    const res = await fetch(url, options);
    if (res.status !== HTTP_CODE.SUCCESS) {
      throw new Error(`Error: ${res.status} ${res.statusText}, ${await res.text()}`);
    }
    const resBody: IPage = await res.json();
    return resBody;
  }
}