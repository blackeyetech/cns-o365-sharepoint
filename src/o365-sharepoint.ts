// Imports here
import CNShell from "cn-shell";
import { CNO365 } from "@cn-shell/o365";

import qs from "qs";

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

  // Public methods here
  async getSiteId(siteName: string): Promise<string> {
    return "";
  }

  async getListId(siteId: string, listName: string): Promise<string> {
    return "";
  }

  async getListItem(siteId: string, listId: string, id: string): Promise<{}> {
    return {};
  }

  async getListItems(
    siteId: string,
    listId: string,
    filter?: {},
  ): Promise<{}[]> {
    return [];
  }

  async updateListItem(
    siteId: string,
    listId: string,
    id: string,
    columns: {},
  ): Promise<boolean> {
    return false;
  }

  async addListItem(
    siteId: string,
    listId: string,
    columns: {},
  ): Promise<boolean> {
    return false;
  }

  async deleteListItem(
    siteId: string,
    listId: string,
    id: string,
  ): Promise<boolean> {
    return false;
  }
}

export { CNO365Sharepoint };
