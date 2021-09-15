// Imports here
import CNShell from "cn-shell";
import { CNO365 } from "@cn-shell/o365";

// import qs from "qs";

import * as MSGraph from "@microsoft/microsoft-graph-types";

// Misc config consts here

// Misc consts here
const GRAPH_API_VERSION = "v1.0";

// process.on("unhandledRejection", error => {
//   // Will print "unhandledRejection err is not defined"
//   console.log("unhandledRejection", error);
// });

// interfaces here

// CNO365Sharepoint class here
class CNO365Sharepoint extends CNO365 {
  // Properties here

  // Constructor here
  constructor(name: string, master?: CNShell) {
    super(name, master);
  }

  // Abstract method implementations here
  async start(): Promise<boolean> {
    await super.start();
    return true;
  }

  async stop(): Promise<void> {
    await super.stop();
    return;
  }

  async healthCheck(): Promise<boolean> {
    return await super.healthCheck();
  }

  // Private methods here
  async getRootSite(): Promise<MSGraph.Site | undefined> {
    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/root`,

      headers: {
        Authorization: `Bearer ${this._token}`,
      },
    }).catch(e => {
      this.error("Error while getting root sharepoint site - (%s)", e);
    });

    if (res === undefined || res.status !== 200) {
      return undefined;
    }

    let site: MSGraph.Site = res.data;

    return site;
  }
  // Public methods here
  async getSiteId(siteName: string): Promise<string | undefined> {
    let rootSite = await this.getRootSite();
    let hostname = rootSite?.siteCollection?.hostname;

    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/${hostname}:/sites/${siteName}`,

      headers: {
        Authorization: `Bearer ${this._token}`,
      },
    }).catch(e => {
      this.error("Error while getting sharepoint sites - (%s)", e);
    });

    if (res === undefined || res.status !== 200) {
      return undefined;
    }

    let site: MSGraph.Site = res.data;

    return site.id;
  }

  async getLists(siteId: string): Promise<MSGraph.List[] | undefined> {
    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/${siteId}/lists`,

      headers: {
        Authorization: `Bearer ${this._token}`,
      },
    }).catch(e => {
      this.error(
        "Error while getting sharepoint lists for site ID (%s) - (%s)",
        siteId,
        e,
      );
    });

    if (res === undefined || res.status !== 200) {
      return undefined;
    }

    let lists: MSGraph.List[] = res.data.value;

    return lists;
  }

  async getListId(
    siteId: string,
    listName: string,
  ): Promise<string | undefined> {
    let lists = await this.getLists(siteId);

    let list = lists?.find(el => el.name === listName);

    if (list === undefined) {
      return undefined;
    }

    return list.id;
  }

  async getListItem(
    siteId: string,
    listId: string,
    id: string,
  ): Promise<MSGraph.ListItem | undefined> {
    let res = await this.httpReq({
      method: "get",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/${siteId}/lists/${listId}/items/${id}?expand=fields`,

      headers: {
        Authorization: `Bearer ${this._token}`,
      },
    }).catch(e => {
      this.error(
        "Error while getting list items for list ID (%s) - (%s)",
        listId,
        e,
      );
    });

    if (res === undefined || res.status !== 200) {
      return undefined;
    }

    return res.data;
  }

  async getListItems(
    siteId: string,
    listId: string,
    select: string[] = [],
    filter?: string,
  ): Promise<MSGraph.ListItem[] | undefined> {
    let url = `${this._resource}/${GRAPH_API_VERSION}/sites/${siteId}/lists/${listId}/items`;

    if (select.length) {
      url = `${url}?expand=fields($select=${select.join(",")})`;
    } else {
      url = `${url}?expand=fields`;
    }

    if (filter !== undefined) {
      url = `${url}&filter=${filter}`;
    }

    let res = await this.httpReq({
      method: "get",
      url,

      headers: {
        Authorization: `Bearer ${this._token}`,
      },
    }).catch(e => {
      this.error(
        "Error while getting list items for list ID (%s) - (%s)",
        listId,
        e,
      );
    });

    if (res === undefined || res.status !== 200) {
      return undefined;
    }

    return res.data.value;
  }

  async updateListItem(
    siteId: string,
    listId: string,
    id: string,
    columns: { [key: string]: any },
  ): Promise<boolean> {
    let res = await this.httpReq({
      method: "patch",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/${siteId}/lists/${listId}/items/${id}/fields`,
      data: columns,
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
    }).catch(e => {
      this.error(
        "Error while updating list item (%s) for list ID (%s) - (%s)",
        id,
        listId,
        e,
      );
    });

    if (res === undefined || res.status !== 200) {
      return false;
    }

    return true;
  }

  async addListItem(
    siteId: string,
    listId: string,
    fields: { [key: string]: any },
  ): Promise<string> {
    let res = await this.httpReq({
      method: "post",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/${siteId}/lists/${listId}/items`,
      data: { fields },
      headers: {
        Authorization: `Bearer ${this._token}`,
        "Content-Type": "application/json",
      },
    }).catch(e => {
      this.error(
        "Error while creating list item for list ID (%s) - (%s)",
        listId,
        e,
      );
    });

    if (res === undefined || res.status !== 201) {
      return "";
    }

    return res.data.id;
  }

  async deleteListItem(
    siteId: string,
    listId: string,
    id: string,
  ): Promise<boolean> {
    let res = await this.httpReq({
      method: "delete",
      url: `${this._resource}/${GRAPH_API_VERSION}/sites/${siteId}/lists/${listId}/items/${id}`,

      headers: {
        Authorization: `Bearer ${this._token}`,
      },
    }).catch(e => {
      this.error(
        "Error while deleteing list item (%s) for list ID (%s) - (%s)",
        id,
        listId,
        e,
      );
    });

    if (res === undefined || res.status !== 204) {
      return false;
    }

    return true;
  }
}

export { CNO365Sharepoint };
