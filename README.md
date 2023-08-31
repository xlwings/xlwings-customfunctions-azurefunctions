# Excel custom functions with xlwings Server on Azure Functions

* Before getting started, have a look at https://github.com/xlwings/xlwings-officejs-quickstart that walks you through setting up a local development environment for Excel custom functions with xlwings Server.

* Also, read the xlwings docs: https://docs.xlwings.org/en/latest/pro/server/officejs_custom_functions.html

This repo shows how to deploy xlwings Server to Azure Functions. We're using the v2 Python programming model using the CLI to deploy the function app. You could, instead, also deploy it via VS Code.

You will benefit by looking at the following resources:

- [Microsoft's Python quickstart guide](https://learn.microsoft.com/en-us/azure/azure-functions/create-first-function-cli-python?tabs=azure-cli%2Cbash&pivots=python-mode-decorators) (this repo was created by following this guide)
- [Microsoft's Python function reference](https://learn.microsoft.com/en-us/azure/azure-functions/functions-reference-python?tabs=asgi%2Capplication-level&pivots=python-mode-decorators).

Note that Azure functions allow you to work with an existing WSGI/ASGI Python web framework, but we're sticking to the native v2 Python programming model here.

## Prerequisites

For the following walk through, you'll need to have the Azure CLI and Azure Functions Core Tools installed, see [here](https://docs.microsoft.com/en-us/cli/azure/install-azure-cli).

Note that deploying this repo can incur costs on your Azure account. You can find the pricing information [here](https://learn.microsoft.com/en-us/azure/azure-functions/create-first-function-cli-python?tabs=azure-cli%2Cbash&pivots=python-mode-decorators#configure-your-local-environment).

If you have an existing workflow to work with Azure functions, you may prefer to stick to that and copy over the relevant parts of this repo.

## Create a function app

While you can run Azure functions locally, we're deploying the function app directly to Azure:

In the commands below, we're going to use the following names/parameters:

- the function app: `xlwings-quickstart`
- the resource group: `xlwings-quickstart-rg`
- the storage account: `xlwingsquickstartsa`
- deploy it to the region: `northeurope`

Note that you may want/need to use different names/parameters though.

Before you begin, you'll need to login to Azure:

```bash
az login
```

1.  Create a resource group:

    ```bash
    az group create --name xlwings-quickstart-rg --location northeurope
    ```

2.  Create storage account:

    ```bash
    az storage account create --name xlwingsquickstartsa --location northeurope --resource-group xlwings-quickstart-rg --sku Standard_LRS
    ```

3.  Create the function app:

    ```bash
    az functionapp create --resource-group xlwings-quickstart-rg --consumption-plan-location northeurope --runtime python --runtime-version 3.10 --functions-version 4 --name xlwings-quickstart --os-type linux --storage-account xlwingsquickstartsa
    ```

4.  Set the xlwings license key as environment variable (you can get a free trial key [here](https://www.xlwings.org/trial)):

    ```bash
    az functionapp config appsettings set --name xlwings-quickstart --resource-group xlwings-quickstart-rg --settings XLWINGS_LICENSE_KEY=<YOUR_LICENSE_KEY>
    ```

    Alternatively, you could also set env vars in the Azure portal, under your Function App's Configuration: `+New application setting`.

5.  Set the following setting to enable the worker process to index the functions:

    ```bash
    az functionapp config appsettings set --name xlwings-quickstart --resource-group xlwings-quickstart-rg --settings AzureWebJobsFeatureFlags=EnableWorkerIndexing
    ```

6.  Deploy the function app (this is also the command to run to deploy an update):

    ```bash
    func azure functionapp publish xlwings-quickstart
    ```

    It should terminate with the following message:

    ```bash
    Remote build succeeded!
    Syncing triggers...
    Functions in xlwings-quickstart:
        custom-functions-call - [httpTrigger]
            Invoke url: https://xlwings-quickstart.azurewebsites.net/api/xlwings/custom-functions-call

            custom-functions-code - [httpTrigger]
                Invoke url: https://xlwings-quickstart.azurewebsites.net/api/xlwings/custom-functions-code

            custom-functions-meta - [httpTrigger]
                Invoke url: https://xlwings-quickstart.azurewebsites.net/api/xlwings/custom-functions-meta

            taskpane - [httpTrigger]
                Invoke url: https://xlwings-quickstart.azurewebsites.net/api/taskpane.html
    ```

    If you don't see the URLs printed, re-run the command. If you still don't see them, there's probably something wrong with your code or you may have a missing package under `requirements.txt`. Check Azure monitoring.

8. On Azure portal, under Function App > Your Function App > CORS, set `Allowed Origins` to `*` if you want to be able to call the functions from Excel on the web. This step should not be required if you're only using the desktop version of Excel.

## Excel add-in (Manifest)

* Replace `https://127.0.0.1:8000` with the URL of your function app in the `manifest.xml` file.
* Sideload the add-in according to your platform, see: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing

When the add-in is loaded, it will show a button on the Home tab. Clicking it will open a task pane. While you won't need the task pane for custom functions, it can be used to display information to the user. Also, after making changes to `taskpane.html`, you can right-click on the task pane and select "Reload" to reload the add-in code.

Note that this repo doesn't come with any icons, so you'll see a default icon in Excel. You can change that by pointing the respective icon URLs in the manifest to your own icons.

## Custom functions

Once the function app is deployed and the add-in is sideloaded, you can play around with the custom functions by using one of the sample functions in `custom_function.py` in a cell, e.g., `=XLWINGS.HELLO("xlwings")`.

## Local development

If you wanted to run the functions locally, add the `local.settings.json` file in the repo's root (it is by default ignored by git):

```javascript
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "python",
    "AzureWebJobsFeatureFlags": "EnableWorkerIndexing",
    "AzureWebJobsStorage": ""
  }
}
```

## Best practices

* Change the `<Id>` in the manifest to a unique GUID for each environment (e.g., dev, prod, etc.).
* Use a different `Functions.Namespace` in the manifest for each environment/version of the add-in. E.g., `DEV`, `MYAPP_V1`, etc. to prevent name clashes with other version of the app.

## Authentication with Azure AD

In `function_app.py`, only `custom_functions_call()` can be authenticated, the rest of the endpoints need to allow for anonymous access as Excel isn't able to load the add-in otherwise.

First, you'll need to create an app in Azure AD and optionally, you can use roles for role-based access control (RBAC). For a walk-through on how to set up Azure AD, see: https://github.com/xlwings/xlwings-server-auth-azuread

* Update your manifest's `WebApplicationInfo` section with `Id` and `Resource` (bottom of the file).
* Create the following env vars in your function app either via Azure portal: `AZUREAD_CLIENT_ID=...`, `AZUREAD_TENANT_ID=...`, `XLWINGS_REQUIRED_ROLES=xlwings.user`, `DB_CONNECTION_STRING=...` or command line:

```bash
az functionapp config appsettings set --name xlwings-quickstart --resource-group xlwings-quickstart-rg --settings AZUREAD_CLIENT_ID=... AZUREAD_TENANT_ID=... XLWINGS_REQUIRED_ROLES=xlwings.user DB_CONNECTION_STRING=...
```

`DB_CONNECTION_STRING` expects a SQLAlchemy connection string pointing to an AdventureWorks sample database.

Since the access token is currently cached in Excel, you'll need to open the side bar of the add-in, right-click and `Reload` to get a new access token after making changes. **Note, however, that changes to Azure AD roles etc. can take 10 minutes or longer to come through!**

## Monitoring / Logs

In the Azure portal, go to your function app, then click on the function name, e.g., `custom-function-call`. Then click on `Monitor` in the left side bar and now on the timestamp of a specific run. Note that it takes multiple minutes for the logs to show up.


## Cleanup

After running this tutorial you can get rid of all the resources again by running:

```bash
az group delete --name xlwings-quickstart-rg
```
