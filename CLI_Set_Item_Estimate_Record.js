/**
 *@NApiVersion 2.x
 *@NScriptType ClientScript
 */
define(['N/record', 'N/currentRecord'],
    function(record) {
       function fieldChanged(context) {
		try{
				var currentRecord = context.currentRecord;
				var fieldObj = context.fieldId;
				var subListObj = context.sublistId;
				var numLines = currentRecord.getLineCount({sublistId: 'item'});
			   
				if(subListObj=="item" && fieldObj=="custcol_item_custom" )
				{
					// alert("field changed");
					var currIndex = currentRecord.getCurrentSublistIndex({sublistId: 'item'});
					//alert("currIndex"+currIndex);
					var currentSubllistItem=currentRecord.getCurrentSublistValue({sublistId:'item',fieldId:'custcol_item_custom', line:currIndex});
				
					if(currentSubllistItem)
					{
						
						currentRecord.setCurrentSublistValue({sublistId: 'item',fieldId: 'item',value: currentSubllistItem,forceSyncSourcing:true});
					}
					
				}
       
			}catch(e){log.error ({title: e.name,details: e.message});}
        }

        return {
            fieldChanged: fieldChanged
        };
    });