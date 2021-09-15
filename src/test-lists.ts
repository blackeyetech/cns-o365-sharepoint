import { CNO365Sharepoint } from "./o365-sharepoint";

import inquirer from "inquirer";
// import EasyTable from "easy-table";

// import fs from "fs";
// import path from "path";

// enums here
enum Prompts {
  TEST = "Test",
  SITE = "Site",
  LIST = "List",
  ITEM = "Item",
}

enum TestChoices {
  UPDATE = "Update",
  DEL = "Delete",
  ITEM = "Item",
  ITEMS = "Items",
  LIST = "List",
  LISTS = "Lists",
  QUIT = "Quit",
}

// Tests here
async function getListId(msSharepoint: CNO365Sharepoint): Promise<void> {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.LIST,
      message: "Input list to use:",
    },
  ]);

  let listName = answer[Prompts.LIST];

  let id = await msSharepoint.getListId(listName);

  console.log(id);
}

async function getLists(msSharepoint: CNO365Sharepoint): Promise<void> {
  let lists = await msSharepoint.getLists();

  if (lists === undefined) {
    return;
  }

  for (let list of lists) {
    console.log(list.id, list.name);
  }
}

async function getItems(msSharepoint: CNO365Sharepoint): Promise<void> {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.LIST,
      message: "Input list to use:",
    },
  ]);

  let listName = answer[Prompts.LIST];
  let listId = await msSharepoint.getListId(listName);

  if (listId === undefined) {
    return;
  }

  let items = await msSharepoint.getListItems(listId);

  if (items === undefined) {
    return;
  }

  for (let item of items) {
    console.log(item.id, item.fields);
  }
}

async function getItem(msSharepoint: CNO365Sharepoint): Promise<void> {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.LIST,
      message: "Input list to use:",
    },
  ]);

  let listName = answer[Prompts.LIST];
  let listId = await msSharepoint.getListId(listName);

  if (listId === undefined) {
    return;
  }

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.ITEM,
      message: "Input item ID to use:",
    },
  ]);

  let itemId = answer[Prompts.ITEM];

  let item = await msSharepoint.getListItem(listId, itemId);

  if (item === undefined) {
    return;
  }

  console.log(item.id, item.fields);
}

async function deleteItem(msSharepoint: CNO365Sharepoint): Promise<void> {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.LIST,
      message: "Input list to use:",
    },
  ]);

  let listName = answer[Prompts.LIST];
  let listId = await msSharepoint.getListId(listName);

  if (listId === undefined) {
    return;
  }

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.ITEM,
      message: "Input item ID to delete:",
    },
  ]);

  let itemId = answer[Prompts.ITEM];

  let success = await msSharepoint.deleteListItem(listId, itemId);

  console.log(success);
}

async function updateItem(msSharepoint: CNO365Sharepoint): Promise<void> {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.LIST,
      message: "Input list to use:",
    },
  ]);

  let listName = answer[Prompts.LIST];
  let listId = await msSharepoint.getListId(listName);

  if (listId === undefined) {
    return;
  }

  answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.ITEM,
      message: "Input item ID to update:",
    },
  ]);

  let itemId = answer[Prompts.ITEM];

  let success = await msSharepoint.updateListItem(listId, itemId, {
    name: "Kieran",
  });

  console.log(success);
}

// Main here
(async () => {
  let answer = await inquirer.prompt([
    {
      type: "input",
      name: Prompts.SITE,
      message: "Input sharepoint site to use:",
    },
  ]);

  let siteName = answer[Prompts.SITE];

  let msSharepoint = new CNO365Sharepoint("Test-Sharepoint", siteName);
  await msSharepoint.init();

  while (1) {
    answer = await inquirer.prompt([
      {
        type: "list",
        name: Prompts.TEST,
        choices: [
          TestChoices.LISTS,
          TestChoices.LIST,
          TestChoices.ITEMS,
          TestChoices.ITEM,
          TestChoices.DEL,
          TestChoices.UPDATE,
          TestChoices.QUIT,
        ],
        message: "What test do you want to run?",
      },
    ]);

    if (answer[Prompts.TEST] === TestChoices.QUIT) {
      break;
    }

    switch (answer[Prompts.TEST]) {
      case TestChoices.LIST:
        await getListId(msSharepoint);
        break;

      case TestChoices.LISTS:
        await getLists(msSharepoint);
        break;

      case TestChoices.ITEMS:
        await getItems(msSharepoint);
        break;
      case TestChoices.ITEM:
        await getItem(msSharepoint);
        break;
      case TestChoices.DEL:
        await deleteItem(msSharepoint);
        break;
      case TestChoices.UPDATE:
        await updateItem(msSharepoint);
        break;
    }
  }

  msSharepoint.exit();
})();
