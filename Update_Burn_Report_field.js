/**
 *@NApiVersion 2.x
 *@NScriptType UserEventScript
 */
define(['N/record'],
    function(record) {
        
        function afterSubmit(context) {
			   try {
				   log.debug('context.type',context.type);
				   // define the mode
					if (context.type !== context.UserEventType.XEDIT && context.type !== context.UserEventType.EDIT)
						return;
					// Get the current record object
					var objRecord = context.newRecord;
					log.debug('objRecord',objRecord);
					var recordId = objRecord.id;
					var objProjectFin = record.load({type:'customrecord_sss_proj_snapshot',id:recordId,isDynamic:false});
					// Fetch the value of project re cprd id to updatw the value
					var projectObjId = objProjectFin.getValue({fieldId:'custrecord_sss_ss_proj'});
					log.debug('projectObjId',projectObjId);
					 if (objProjectFin.getValue('custrecord_sss_ss_proj')) 
						{
						var burnReportNote = objRecord.getValue({fieldId:'custrecord_burn_report_update'});
						log.debug('burnReportNote',burnReportNote);
						// load the project record
						var ObjProject = record.load({
								type: 'job',
								id:projectObjId,
								isDynamic: true
							});//xedit 
							
							log.debug('ObjProject',ObjProject);
							// set the burn report note field form SSS Project Financial Snapshot record 
							ObjProject.setValue({fieldId:'custentity_prj_brnreportnote',value:burnReportNote});
							
							var projectId = ObjProject.save();// save the project record
							log.debug('project record updaated successfully', 'Id: ' + projectId);
						}
				   }
					catch (e) {
						log.error(e.name);
					}
            
        }
        return {
          
            afterSubmit: afterSubmit
        };
    });