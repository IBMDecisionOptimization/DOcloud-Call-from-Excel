
# IBM DOcplexcloud Call from Excel Sample

This sample shows how you can leverage the power of the [DOcplexcloud service](https://dropsolve-oaas.docloud.ibmcloud.com/software/analytics/docloud) directly from your Microsoft Excel workbook. 


## Prerequisites

1. You need Microsoft Excel 2010 or higher installed.
2. You need to be registered to [IBM Decision Optimization on Cloud](https://onboarding-oaas.DOcloud.ibmcloud.com/software/analytics/docloud/) (free trial available).
3. You need the [IBM DOcplexcloud base URL and an API key](https://onboarding-oaas.docloud.ibmcloud.com/software/analytics/docloud/).

## Configuration

In the worksheet called **Dashboard**, you have to:
 * specify the base URL.
 * specify your API key.

## About this sample

This sample contains a Microsoft Excel Workbook containing VBA macros.  
The workbook allows you to solve the OPL model defined in the worksheet **Model** using the data defined in the other worksheets. [See here](https://developer.ibm.com/docloud/docs/more-information/opl-model-input-and-output/) for details on Excel as input format.

The optimization model in this sample is the *Factory planning* example. You can find more details on this model in the sample available in the DropSolve [FAQ & Samples](https://dropsolve-oaas.docloud.ibmcloud.com/dropsolve/doc).

To see the code used by this sample you can use Microsoft Visual Basic for Applications. To open this, use **Alt+F11** in Excel.  
The code is in the DOcplexcloud module. To interact with the DOcplexcloud service, we use the DOcplexcloud REST API. For more information, you can access the [REST API Reference](https://developer.ibm.com/docloud/docs/rest-api/rest-api-documentation/) in the Developer Centre.

For example, the code to create a new job in DOcplexcloud is:

	Dim docloudService As New WinHttpRequest
    With docloudService
    	.Open "POST", https://xxxx.ibmcloud.com/yyyy/rest/v1/jobs, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "X-IBM-Client-Id", api_XXXXXX-XXXX-XXXX-XXXX-XXXXXXXX
        .send "{ ""parameters"" : <list of parameters> }"
    End With

When you click the **Solve** button, it creates and solves your problem on DOcplexcloud. When a result is available, values for output elements are imported into the worksheet **Results.plan**. These data generate the **Report** worksheet, containing a Pivot Chart showing the new production plan per product per month.

If you see these messages:
  * `Subscription [ODSTRIAL:XXXX] of user api_1111-11111-1111 has a limit of 5 jobs total`. In this case, you have to connect to [DropSolve](https://dropsolve-oaas.docloud.ibmcloud.com/dropsolve) and either remove one problem or check the option in dashboard worksheet to delete all existing jobs from DropSolve.
  * `Still running (you can increase the number of retry)`. In this case, the loop that waits for the results has been waiting to long.
You can increase the value for the 'Nb retry' parameter. The loop waits 1 second before a retry.

## How to reuse this Workbook with your own model
1. Replace the model contained in the tab **Model** (content of Cell **A1**).
2. Delete worksheets containing data used by the previous model (all worksheets except **Dashboard** and **Model**).
3. For each input element of your model, you have to create a new tab containing the data (see [Excel as input format](https://developer.ibm.com/docloud/docs/more-information/opl-model-input-and-output/) for more details).
4. After the solve, a new sheet ( **Results.<name of the element>** ) is created for each output element of your model.
5. You can also adapt the code that generates the pivot chart (see the procedure `CreateChart` in the code). Otherwise, the worksheet containing the chart remains empty.

## License

This sample is delivered under the Apache License Version 2.0, January 2004 (see LICENSE.txt). 