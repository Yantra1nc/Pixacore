/**
 * @NApiVersion 2.x
 * @NScriptType ClientScript
*/
define(['N/currentRecord','N/search'], function(currentRecord,search) {
	
	var globalObj	= '';
	
	function pageInit(context) {
		globalObj	= context.currentRecord;
		//alert("globalObj"+globalObj);
		
		var customerId = globalObj.getValue({fieldId:'custpage_registered_person'});
		var year_id = globalObj.getValue({fieldId:'custpage_year_range'});
		var resultIndexSN	= 0; 
		var resultStepSN	= 1000;
		var invoiceFilter	= [];
			var invoiceColumn	= [];
		if(customerId) {
					invoiceFilter.push(search.createFilter({name: "customer", operator: search.Operator.ANYOF, values: customerId}));
				}
			
			invoiceFilter.push(search.createFilter({name: "custentity_mp_client_po_num", operator: search.Operator.ISNOTEMPTY, values: ''}));	
			invoiceColumn.push(search.createColumn({ name: "custentity_mp_client_po_num",summary: "GROUP",label: "Client PO Number"}));
			var searchObjCrdMemo = search.create({type:"job",filters: invoiceFilter, columns: invoiceColumn});
			var searchCount = searchObjCrdMemo.runPaged().count;
			//alert("searchCount"+searchCount)
			log.debug('searchCountCDNR',searchObjCrdMemo.runPaged().count);
			var vendfield = globalObj.getField('custpage_gstin');
           vendfield.removeSelectOption({value : null});
			 if(searchCount != 0)
					{
						do
						{
							var searchResultSN = searchObjCrdMemo.run().getRange({start: resultIndexSN, end: resultIndexSN + resultStepSN});
							if(searchResultSN.length > 0)
							{
								for(var s in searchResultSN )
								{
									vendfield.insertSelectOption({
								value:searchResultSN[s].getValue({ name: "custentity_mp_client_po_num",summary: "GROUP",label: "Client PO Number"}),
								text:searchResultSN[s].getValue({ name: "custentity_mp_client_po_num",summary: "GROUP",label: "Client PO Number"})
									})
								}
							}
							 resultIndexSN = resultIndexSN + resultStepSN;
						} while (searchResultSN.length > 0);
					}
	}
	
	function getFieldData()
	{
		//alert("Entered");
		var monthId	= globalObj.getValue({fieldId: "custpage_month_range"});
   //   alert("monthId"+monthId)
		var yearId = globalObj.getValue({fieldId: "custpage_year_range"});
     //  alert("yearId"+yearId)
		var cust_gstin = globalObj.getValue({fieldId: "custpage_gstin"});
     // alert("cust_gstin"+cust_gstin)
		var customerId = globalObj.getValue({fieldId: "custpage_registered_person"});
      var custName = globalObj.getText({fieldId: "custpage_registered_person"});
		// alert("custName"+custName)
		var baseUrl = window.location.href;
		
		baseUrl = baseUrl.substring(0, baseUrl.indexOf("deploy=1")+8);
		if(monthId) {
			baseUrl=baseUrl+"&monthid="+monthId
		}
		if(yearId) {
			baseUrl=baseUrl+"&yearid="+yearId
		}
		if(cust_gstin) {
			baseUrl=baseUrl+"&cust_gstin="+cust_gstin
		}
		if(customerId) {
			baseUrl=baseUrl+"&customerid="+customerId
		}
      if(custName)
        {
          baseUrl=baseUrl+"&customername="+custName
        }
    //  alert("baseUrl"+baseUrl)
		window.onbeforeunload = null;
		window.location.href = baseUrl;
	}
	function fieldChanged(context)
	{
		var fieldName = context.fieldId
		var sublistName = context.sublistId
		var customerId = globalObj.getValue({fieldId:'custpage_registered_person'});
		var year_id = globalObj.getValue({fieldId:'custpage_year_range'});
		//alert("year_id",year_id)
		var resultIndexSN	= 0; 
		var resultStepSN	= 1000;
		if(fieldName=="custpage_year_range" ||fieldName=="custpage_registered_person" )
		{
			var invoiceFilter	= [];
			var invoiceColumn	= [];
			if(customerId) {
					invoiceFilter.push(search.createFilter({name: "customer", operator: search.Operator.ANYOF, values: customerId}));
				}
			/*if(year_id) {
				invoiceFilter.push(search.createFilter({name: "custentity_pixacore_project_year", operator: search.Operator.IS, values: year_id}));
			}*/
			invoiceFilter.push(search.createFilter({name: "custentity_mp_client_po_num", operator: search.Operator.ISNOTEMPTY, values: ''}));	
			invoiceColumn.push(search.createColumn({ name: "custentity_mp_client_po_num",summary: "GROUP",label: "Client PO Number"}));
			var searchObjCrdMemo = search.create({type:"job",filters: invoiceFilter, columns: invoiceColumn});
			var searchCount = searchObjCrdMemo.runPaged().count;
			//alert("searchCount"+searchCount)
			log.debug('searchCountCDNR',searchObjCrdMemo.runPaged().count);
			var vendfield = globalObj.getField('custpage_gstin');
           vendfield.removeSelectOption({value : null});
			 if(searchCount != 0)
					{
						do
						{
							var searchResultSN = searchObjCrdMemo.run().getRange({start: resultIndexSN, end: resultIndexSN + resultStepSN});
							if(searchResultSN.length > 0)
							{
								for(var s in searchResultSN )
								{
									vendfield.insertSelectOption({
								value:searchResultSN[s].getValue({ name: "custentity_mp_client_po_num",summary: "GROUP",label: "Client PO Number"}),
								text:searchResultSN[s].getValue({ name: "custentity_mp_client_po_num",summary: "GROUP",label: "Client PO Number"})
									})
								}
							}
							 resultIndexSN = resultIndexSN + resultStepSN;
						} while (searchResultSN.length > 0);
					}
				
		/*	var vendfield = globalObj.getField('custpage_multi_select');
           vendfield.removeSelectOption({value : null});
			var jobSearchObj = search.create({
			type: "job",
			filters:
			[
			["custentity_mp_client_po_num","isnotempty",""]
			],
			columns:
			[
			search.createColumn({
         name: "custentity_mp_client_po_num",
         summary: "GROUP",
         label: "Client PO Number"
      })
			]
			});
		var searchResultCount = jobSearchObj.runPaged().count;
		log.debug("jobSearchObj result count",searchResultCount);
		jobSearchObj.run().each(function(result){
			vendfield.insertSelectOption({
				value:result.getValue({ name: "custentity_mp_client_po_num",
         summary: "GROUP",
         label: "Client PO Number"}),
				text:result.getValue({ name: "custentity_mp_client_po_num",
         summary: "GROUP",
         label: "Client PO Number"})
			})
		   // .run().each has a limit of 4,000 results
		   return true;
		});*/
		}
	}
	
	return {
		pageInit : pageInit,
		getFieldData: getFieldData,
		fieldChanged:fieldChanged
	}
});