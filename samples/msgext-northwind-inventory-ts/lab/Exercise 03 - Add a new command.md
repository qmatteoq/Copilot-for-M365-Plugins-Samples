# Building Message Extensions for Microsoft Copilot for Microsoft 365

TABLE OF CONTENTS

* [Welcome](./Exercise%2000%20-%20Welcome.md) 
* [Exercise 1](./Exercise%2001%20-%20Set%20up.md) - Set up your development Environment
* [Exercise 2](./Exercise%2003%20-%20Run%20in%20Copilot.md) - Run the sample as a Copilot plugin 
* [Exercise 3](./Exercise%2003%20-%20Add%20a%20new%20command.md) - Add a new command **(THIS PAGE)**
* [Optional - Code Tour](./Optional%20-%20Code%20tour.md) - Code tour
* [Optional - Message Extension](./Optional%20-%20Run%20sample%20app.md) - Run the sample as a Message Extension

## Exercise 3 - Add a new command 

In this exercise, you will enhance the Teams Message Extension / Copilot plugin  by adding a new command. While the current message extension effectively provides information about products within the Northwind inventory database, it does not provide information related to Northwind’s customers. Your task is to introduce a new command associated with an API call that retrieves products ordered by a customer name specified by the user.

Do this we will:

### 1. Extend the Message Extension User Interface:

We'll do this by updating the manifest.json file to include an option for querying for products by company name. 

### 2. Manifest Update:

In the manifest.json file, we'll provide clear and concise descriptions for the new command and its associated parameters. These descriptions are crucial because Copilot will use them to determine user intents that align with the functionality of this command.

File update: manifest.json

### 3. Message Routing:
We'll extend the method that is called by the Bot Framework when the user queries the Northwind database (**handleTeamsMessagingExtensionQuery**). We'll add the call to our message handler that's invoked when the user searches for products by Company name. 

File update: searchApp.ts

```text
Note that in the UI-based operation of the Message Extension / plugin, this command is explicitly called. However, when invoked by Microsoft 365 Copilot, the command is triggered by the Copilot orchestrator.
```

### 4. Implement the message handler:

We will implement the search for products by Company name and return a list of products to be displayed in the Message Extension or in the Copilot chat UI.

New file: 
```
.\src\messageExtensions\customerSearchCommand.ts
```

### Update manifest.json

1. Open manifest.json and add the following json immediately after the **discountSearch** command. 

```json
{
        "id": "companySearch",
        "context": [
            "compose",
            "commandBox"
        ],
        "description": "Given a company name, Search for products ordered by that company",
        "title": "Customer",
        "type": "query",
        "parameters": [
            {
                "name": "companyName",
                "title": "Company name",
                "description": "The company name to find products ordered by that company",
                "inputType": "text"
            }
        ]
    }
```
Note: 

The "id" is the connection between the UI and the code. This value is defined as COMMAND_ID in the *SearchCommand.ts files. See how each of these files has a unique COMMAND_ID that corresponds to the value of "id".


### Implement product search by Company

*We will first implement the message handler and then add the routing to it when completed.*

We will use a lot of the code created for the other handlers. 

1. In VS Code copy 'productSearchCommand.ts' and paste in the same folder to create a copy. Rename this file **customerSearchCommand.ts**.





