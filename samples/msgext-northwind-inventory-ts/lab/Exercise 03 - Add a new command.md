# Building Message Extensions for Microsoft Copilot for Microsoft 365

TABLE OF CONTENTS

* [Welcome](./Exercise%2000%20-%20Welcome.md) 
* [Exercise 1](./Exercise%2001%20-%20Set%20up.md) - Set up your development Environment
* [Exercise 2](./Exercise%2003%20-%20Run%20in%20Copilot.md) - Run the sample as a Copilot plugin 
* [Exercise 3](./Exercise%2003%20-%20Add%20a%20new%20command.md) - Add a new command **(THIS PAGE)**
* [Optional - Code Tour](./Optional%20-%20Code%20tour.md) - Code tour
* [Optional - Message Extension](./Optional%20-%20Run%20sample%20app.md) - Run the sample as a Message Extension

## Exercise 3 - Add a new command 

In this exercise, you will enhance the Teams Message Extension (ME) / Copilot plugin  by adding a new command. While the current message extension effectively provides information about products within the Northwind inventory database, it does not provide information related to Northwind’s customers. Your task is to introduce a new command associated with an API call that retrieves products ordered by a customer name specified by the user. This exercise assumes you have completed exercises 1 and 2.

To do this you'll complete five tasks.
1. **Extend the Message Extension / plugin User Interface** by modifying the Teams app manifest. This includes introducing a new command: **"companySearch"**. Note the UI for the Message Extension is an adaptive card where for Copilot it is text input and output in Copilot chat.
2. **Create a handler for the 'companySearch' command**. This will parse the query string passed in from the message routing code, validate the input and call the product search by company API. This step will also populate an adaptive card with the returned product list which will be displayed in the ME or Copilot chat UI.
3. Update the command **routing** code to route the new command to the handler created in the previous step. You'll do this by extending the method called by the Bot Framework when users query the Northwind database (**handleTeamsMessagingExtensionQuery**). 
4. **Implement Product Search by Company** that returns a list of products ordered by that company.
5. **Run the app** and search of products that were purchased by a specified company.

# Step 1 - Extend the Message Extension / plugin User Interface 

1. Open **manifest.json** and add the following json immediately after the **discountSearch** command. Here you're adding to the **"commands"** array which defines the list of commands supported by the ME / plugin.

```json
{
        "id": "companySearch",
        "context": [
            "compose",
            "commandBox"
        ],
        "description": "Given a company name, search for products ordered by that company",
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
```
Note: The "id" is the connection between the UI and the code. This value is defined as COMMAND_ID in the *SearchCommand.ts files. See how each of these files has a unique COMMAND_ID that corresponds to the value of "id".
```

# Step - 2 Create a handler for the 'companySearch' command
We will use a lot of the code created for the other handlers. 

1. In VS Code copy '**productSearchCommand.ts**' and paste into the same folder to create a copy. Rename this file **customerSearchCommand.ts**.

2. Change line 10 to:
```javascript
const COMMAND_ID = "companySearch";
```

2. Replace the content of **handleTeamsMessagingExtensionQuery** with:
```javascript
 {
    let companyName;

    // Validate the incoming query, making sure it's the 'companySearch' command
    // The value of the 'companyName' parameter is the company name to search for
    if (query.parameters.length === 1 && query.parameters[0]?.name === "companyName") {
        [companyName] = (query.parameters[0]?.value.split(','));
    } else { 
        companyName = cleanupParam(query.parameters.find((element) => element.name === "companyName")?.value);
    }
    console.log(`🍽️ Query #${++queryCount}:\ncompanyName=${companyName}`);    

    const products = await searchProductsByCustomer(companyName);

    console.log(`Found ${products.length} products in the Northwind database`)
    const attachments = [];
    products.forEach((product) => {
        const preview = CardFactory.heroCard(product.ProductName,
            `Customer: ${companyName}`, [product.ImageUrl]);

        const resultCard = cardHandler.getEditCard(product);
        const attachment = { ...resultCard, preview };
        attachments.push(attachment);
    });
    return {
        composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments: attachments,
        },
    };
}
```
Note that you will implement searchProductsByCustomer in the next step.

# Step - 3 Update the command routing
In this step you will route the '**companySearch**' command to the handler you implemented in the previous step.

2. Open **searchApp.ts** and add the following. Underneath this line:
```javascript
import revenueSearchCommand from "./messageExtensions/revenueSearchCommand";
```
Add this line:
```javascript
import customerSearchCommand from "./messageExtensions/customerSearchCommand";
```

3. Underneath this statement:
```javascript
      case revenueSearchCommand.COMMAND_ID: {
        return revenueSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
      }
```
Add this statement:
```javascript
      case customerSearchCommand.COMMAND_ID: {
        return customerSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
      }
```
```text
Note that in the UI-based operation of the Message Extension / plugin, this command is explicitly called. However, when invoked by Microsoft 365 Copilot, the command is triggered by the Copilot orchestrator.
```
# Step - 4 Implement Product Search by Company
 We will implement the search for products by Company name and return a list of products ordered by the company. 

1. Open **.\src\northwindDB\products.ts**

2. Update the 'import' statement on line 1 to include Order and Customer. It should look as follows
```javascript
import {
    TABLE_NAME, Product, ProductEx, Supplier, Category, OrderDetail,
    Order, Customer
} from './model';
```
3. Add the new function **searchProductsByCustomer**

Underneath this line:
```javascript
import { getInventoryStatus } from '../adaptiveCards/utils';
```
add the function:
```javascript
export async function searchProductsByCustomer(companyName: string): Promise<ProductEx[]> {

    let result = await getAllProductsEx();

    let customers = await loadReferenceData<Customer>(TABLE_NAME.CUSTOMER);
    let customerId="";
    for (const c in customers) {
        if (customers[c].CompanyName.toLowerCase() === companyName.toLowerCase()) {
            customerId = customers[c].CustomerID;
            break;
        }
    }
    
    if (customerId === "") 
        return [];

    let orders = await loadReferenceData<Order>(TABLE_NAME.ORDER);
    let orderdetails = await loadReferenceData<OrderDetail>(TABLE_NAME.ORDER_DETAIL);
    // build an array orders by customer id
    let customerOrders = [];
    for (const o in orders) {
        if (customerId === orders[o].CustomerID) {
            customerOrders.push(orders[o]);
        }
    }
    
    let customerOrdersDetails = [];
    // build an array order details customerOrders array
    for (const od in orderdetails) {
        for (const co in customerOrders) {
            if (customerOrders[co].OrderID === orderdetails[od].OrderID) {
                customerOrdersDetails.push(orderdetails[od]);
            }
        }
    }

    // Filter products by the ProductID in the customerOrdersDetails array
    result = result.filter(product => 
        customerOrdersDetails.some(order => order.ProductID === product.ProductID)
    );

    return result;
}
```
# Step 5 - Run the App! Search for product by company name

1. Delete the 'Northwest Inventory' app in Teams. This step is necessary since you are updating the manifest. Manifest updates require the app to be reinstalled. The cleanest way to do this is to first delete it in Teams.

    a. In Teams, right click on '**Apps**' in the left rail, search for 'Northwest Inventory'. 

    b. Right click on the 'Northwest Inventory' icon and select uninstall.

2. Like you did in [Exercise 2](./Exercise%2003%20-%20Run%20in%20Copilot.md) - Run the sample as a Copilot plugin, start the app in Visual Studio Code using teh  'Debug in Teams (Edge)' profile.

3. In Teams click on 'Chat' and then 'Copilot'. Copilot should be the top-most option.
4. Click on the 'Plugin icon' and select 'Northwind Inventory' to enable the plugin.
5. Enter the prompt: 
```
What are the products ordered by 'Consolidated Holdings' in Northwind Inventory?
```




