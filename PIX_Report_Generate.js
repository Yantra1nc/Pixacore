//DATE - 29-04-2020
/**
 *@NApiVersion 2.0
 *@NScriptType Suitelet
 **/
define(['N/ui/serverWidget', 'N/runtime', 'N/search', 'N/encode', 'N/file', 'N/record', 'N/format', 'N/config'], function(serverWidget, runtime, search, encode, file, record, format, config) {
	 var jsonAyyaData=[];
	function onRequest(context)
	{
		var reqObj	= context.request;
		try {
			if(reqObj.method == "GET") {
				var monthId = reqObj.parameters.monthid;
					log.debug({title: " monthId : ",details: monthId});
					var yearId = reqObj.parameters.yearid;
					
					log.debug({title: " yearId : ",details: yearId});
					var gstIn = reqObj.parameters.cust_gstin;
					log.debug({title: " gstIn : ",details: gstIn});
					var gstCustomerId = reqObj.parameters.customerid;
					log.debug('gstCustomerId',gstCustomerId)
					var customer_name = reqObj.parameters.customername;
					log.debug('customer_name',customer_name)
					
				var form= serverWidget.createForm({title: "Monthly Fee Reconciliation"});
				form.clientScriptModulePath = '/SuiteScripts/Return_Suitelet_URL_gstr3B_report.js';
				
				var searchButton		= form.addButton({id: 'custpage_search', label: "Search", functionName: "getFieldData()"});
				
				var customerField = form.addField({
							id: 'custpage_registered_person',
							label: "Customer Name",
							type: serverWidget.FieldType.SELECT,
							source: 'customer'
						});
						var gstin = form.addField({
							id: 'custpage_gstin',
							label: "Client Po Number",
							type: serverWidget.FieldType.MULTISELECT
						});
						/*gstin.addSelectOption({
							value:'',
							text:''
						})*/
					//	clientPoNumner(gstin);
						/*gstin.updateDisplayType({
							displayType: serverWidget.FieldDisplayType.HIDDEN
						});*/
						var monthRange = form.addField({
							id: 'custpage_month_range',
							label: "Month",
							type: serverWidget.FieldType.SELECT
						});
					/*	monthRange.updateDisplayType({
							displayType: serverWidget.FieldDisplayType.HIDDEN
						});*/
						var yearRange = form.addField({
							id: 'custpage_year_range',
							label: "Year",
							type: serverWidget.FieldType.MULTISELECT
						});
						setMonthYearData(monthRange, yearRange);
					/*	var multiSelect = form.addField({
							id: 'custpage_multi_select',
							label: "multiline PO#",
							type: serverWidget.FieldType.MULTISELECT,
							//source:'subsidiary'
						});
						clientPoNumner(multiSelect)
						
				*/
				var emp_text = form.addField({
							id: 'custpage_emp_text',
							label: "Emp Name",
							type: serverWidget.FieldType.TEXT
						});
						emp_text.updateDisplayType({
							displayType: serverWidget.FieldDisplayType.HIDDEN
						});
				var exportButton		= form.addSubmitButton({id: "custpage_sub_butt", label:"Export"});
				//var tdsSecwise	= form.addField({id: "custpage_sec_wise", label: "TDS Section Wise", type: serverWidget.Type.});
				
				var tdsReport			= form.addSubtab({id: 'custpage_tds_sublist', label: 'MONTHLY REPORT'});
				var htmlFile			= form.addField({id: 'custpage_html', label: 'Export', type: serverWidget.FieldType.INLINEHTML, container: 'custpage_tds_sublist'});
				var excelFile			= form.addField({id:'custpage_excel', label: 'Print', type: serverWidget.FieldType.INLINEHTML, container: 'custpage_tds_sublist'});
				
				//log.debug({title: "panObj", details:panObj});
				
			/*	var params = {};
						if(monthId) {
							params.monthid = monthId;
						}
						if(yearId) {
							params.yearid = yearId;
						}
						if(gstIn) {
							params.gstin = gstIn;
						}
						if(gstCustomerId) {
							params.gstcustomerid = gstCustomerId;
						}
						*/
				if(yearId) {
					yearRange.defaultValue= yearId;
				}
				if(gstIn) {
					gstin.defaultValue= gstIn;
				}
				if(gstCustomerId) {
					customerField.defaultValue= gstCustomerId;
				}
				if(monthId) {
					monthRange.defaultValue= monthId;
				}
				if(customer_name)
				{
					emp_text.defaultValue=customer_name
				}
				
				var tdsReportData	= tds_report_data(gstIn,yearId,gstCustomerId,monthId,customer_name);
				var bothData		= tdsReportData.split(":||:");
				var htmlData 		= bothData[0];
				var excelData 		= bothData[1];
				//log.debug({title: "htmlData", details: htmlData});
				//log.debug({title: "excelData", details: excelData});
				
				htmlFile.defaultValue = htmlData;
				excelFile.defaultValue= excelData;
				excelFile.updateDisplayType({displayType: serverWidget.FieldDisplayType.HIDDEN});
				context.response.writePage({pageObject: form});
			}
			else {
				var excelFile	= reqObj.parameters['custpage_excel'];
				log.debug({title: "excelFile", details: excelFile});
				var custName = reqObj.parameters['custpage_registered_person'];
				log.debug('custName',custName);
				var newName = reqObj.parameters.custpage_emp_text;
				log.debug('newName',newName)
				var yearValue = reqObj.parameters['custpage_year_range'];
				log.debug('yearValue',yearValue);
				var crrMonth = reqObj.parameters['custpage_month_range'];
				log.debug('crrMonth',crrMonth);
				var dateValue= new Date();
				var monthValue = dateValue.getMonth()
				//var mulValue = reqObj.parameters['custpage_multi_select'];
			//	var mum_value = mulValue.split("");
			//	log.debug('mulValue',mum_value)
			const monthNames = ["January", "February", "March", "April", "May", "June",
				"July", "August", "September", "October", "November", "December"
			];
			var month_Name=''
				if(crrMonth)
					{
						month_Name = monthNames[crrMonth]
					}else{
						month_Name = monthNames[monthValue]
						
					}
				var c_name='';
				
				
				// var recObj = record.load({type:'customer', id:custName, isDynamic:true});
				// var c_name = recObj.getText({fieldId:'entityid'});
				var xmlStr = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
				xmlStr += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ';
				xmlStr += 'xmlns:o="urn:schemas-microsoft-com:office:office" ';
				xmlStr += 'xmlns:x="urn:schemas-microsoft-com:office:excel" ';
				xmlStr += 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ';
				xmlStr += 'xmlns:htmlObj1="http://www.w3.org/TR/REC-html40">';

				xmlStr += '<Worksheet ss:Name="Report">';
				
				xmlStr += '<Table>';

					xmlStr += '<Row>' + '<Cell ss:MergeAcross="5"><Data ss:Type="String">'+newName +'-'+ month_Name.substring(0,3)+' Fee Reconciliation - '+yearValue+'</Data></Cell>' +
						'<Cell ss:MergeAcross="2" ><Data ss:Type="String" > Project SOW Estimate</Data></Cell>'+
						'<Cell ss:MergeAcross="2"><Data ss:Type="String">2021 Brief Estimate</Data></Cell>' +
						'<Cell ><Data ss:Type="String">Variance</Data></Cell>' +
						'<Cell ss:MergeAcross="2"><Data ss:Type="String">Project Actuals</Data></Cell>' +
						'<Cell ><Data ss:Type="String">Variance</Data></Cell>' +
						'<Cell ><Data ss:Type="String">YTD Billing(fees and OOP)</Data></Cell>' +
						'<Cell ><Data ss:Type="String"></Data></Cell>' +
						'<Cell ss:MergeAcross="11"><Data ss:Type="String">Actual Fees and Forecast</Data></Cell>' +
						'<Cell ss:MergeAcross="11"><Data ss:Type="String">Actual OOPs and Forecast</Data></Cell></Row>';

					xmlStr += '<Row>' + '<Cell><Data ss:Type="String">Client PO Number</Data></Cell>' +
						'<Cell><Data ss:Type="String">Project Number</Data></Cell>' +
						'<Cell><Data ss:Type="String">Client : Project</Data></Cell>' +
						'<Cell><Data ss:Type="String">Status</Data></Cell>' +
						'<Cell><Data ss:Type="String">Start Date</Data></Cell>' +
						'<Cell><Data ss:Type="String">Projected End Date</Data></Cell>' +
						/*add on*/
						'<Cell><Data ss:Type="String">Fee $</Data></Cell>' +
						'<Cell><Data ss:Type="String">OOP $</Data></Cell>' +
						'<Cell><Data ss:Type="String">Total Fee and OOP $</Data></Cell>' +
						'<Cell><Data ss:Type="String">Estimated Fees</Data></Cell>' +
						'<Cell><Data ss:Type="String">Estimated OOP</Data></Cell>' +
						'<Cell><Data ss:Type="String">Estimate Total</Data></Cell>' +
						'<Cell><Data ss:Type="String">(Over)/Under</Data></Cell>' +
						'<Cell><Data ss:Type="String">Actual Fee</Data></Cell>' +
						'<Cell><Data ss:Type="String">Actual OOP</Data></Cell>' +
						'<Cell><Data ss:Type="String">Actual Total</Data></Cell>' +
						'<Cell><Data ss:Type="String">(Over)/Under</Data></Cell>' +
						'<Cell><Data ss:Type="String">% Utilized</Data></Cell>' +
						'<Cell><Data ss:Type="String">Billed to Date</Data></Cell>' +
						'<Cell><Data ss:Type="String">Jan</Data></Cell>' +
						'<Cell><Data ss:Type="String">Feb</Data></Cell>' +
						'<Cell><Data ss:Type="String">Mar</Data></Cell>' +
						'<Cell><Data ss:Type="String">Apr</Data></Cell>' +
						'<Cell><Data ss:Type="String">May</Data></Cell>' +
						'<Cell><Data ss:Type="String">June</Data></Cell>' +
						'<Cell><Data ss:Type="String">July</Data></Cell>' +
						'<Cell><Data ss:Type="String">Aug</Data></Cell>' +
						'<Cell><Data ss:Type="String">Sep</Data></Cell>' +
						'<Cell><Data ss:Type="String">Oct</Data></Cell>' +
						'<Cell><Data ss:Type="String">Nov</Data></Cell>' +
						'<Cell><Data ss:Type="String">Dec</Data></Cell>' +
						'<Cell><Data ss:Type="String">Jan</Data></Cell>' +
						'<Cell><Data ss:Type="String">Feb</Data></Cell>' +
						'<Cell><Data ss:Type="String">Mar</Data></Cell>' +
						'<Cell><Data ss:Type="String">Apr</Data></Cell>' +
						'<Cell><Data ss:Type="String">May</Data></Cell>' +
						'<Cell><Data ss:Type="String">Jun</Data></Cell>' +
						'<Cell><Data ss:Type="String">July</Data></Cell>' +
						'<Cell><Data ss:Type="String">Aug</Data></Cell>' +
						'<Cell><Data ss:Type="String">Sep</Data></Cell>' +
						'<Cell><Data ss:Type="String">Oct</Data></Cell>' +
						'<Cell><Data ss:Type="String">Nov</Data></Cell>' +
						'<Cell><Data ss:Type="String">Dec</Data></Cell></Row>';	
				xmlStr += excelFile;

				xmlStr += '</Table></Worksheet>';
				xmlStr += '</Workbook>';
				var fileName		= "MONTHLY_Report"+new Date()+".xls";

				var encodedString	= encode.convert({string: xmlStr, inputEncoding: encode.Encoding.UTF_8, outputEncoding: encode.Encoding.BASE_64});
				var fileObj			= file.create({name: fileName, fileType: file.Type.EXCEL,contents: encodedString});
				context.response.writeFile({file: fileObj});
				

			}
		}
		catch(exp) {
			log.error({title:"Exception log", details: exp.id});
			log.error({title:"Exception log", details: exp.message});
		}
		
	}
	
	function tds_report_data(gstIn,yearId,gstCustomerId,monthId,customer_name) 
	{
		var htmlObj1		= "";
		var excelObj 		= "";
		
			var date = new Date();
			var month = date.getMonth() -1;
			var month_Name = '';
			const monthNames = ["January", "February", "March", "April", "May", "June",
				"July", "August", "September", "October", "November", "December"
			];
			if(yearId)
				
				{
					yearId=yearId;
				}else{
					yearId=''
				}
				
				 var totalValue = 0;
					var c1 =0;
					var c2 =0;
					var c3 =0;
					var c4 =0;
					var c5 =0;
					var c6 =0;
					var c7 =0;
					var c8 =0;
					var c9 =0;
					var c10=0;
					var c11=0;
					var c12=0;
					var c13=0;
					var c14=0;
					var c15=0;
					var c16=0;
					var c17=0;
					var c18=0;
					var c19=0;
					var c20=0;
					var c21=0;
					var c22=0;
					var c23=0;
					var c24=0;
					var c25=0;
					var c26=0;
					var c27=0;
					var c28=0;
					var c29=0;
					var c30=0;
					var c31=0;
					var c32=0;
					var c33=0;
					var c34=0;
					var c36=0;
					var c37=0;
					var c38=0;
					var subTotalCrrMonth=0;
					var subTotalOppMonth=0;
					
					var Ja    =0
					var Fe    =1
					var Mr    =2
					var Ap    =3
					var Ma    =4
					var Ju    =5
					var Jl    =6
					var Au    =7
					var Se    =8
					var Oc    =9
					var No    =10
					var De    =11
				 
			var monthName = monthNames[date.getMonth() + 1];
			log.debug("The current month is ", monthNames[date.getMonth()]);
			var m_Date = date.getMonth()+1;
			// current month -1
			//log.debug("The current month-1 is ", monthNames[monthId-1]);
			
			if(monthId)
			{
				month_Name = monthNames[monthId]
			}else{
				month_Name = monthNames[month]
				
			}
			var c_name='';
			if(gstCustomerId)
			{
			var recObj = record.load({type:'customer', id:gstCustomerId, isDynamic:true});
		    c_name = recObj.getText({fieldId:'entityid'});
			}
			var multiValue = ["US-11066-1","US-11067-1","US-11252-1"]
		var htmlObj1  ='';
		var excelObj ='';
		htmlObj1 += '<table class="minimalistBlack" style="border: 1px solid #000000;width: 100%;text-align: left;  border-collapse: collapse;">';
			htmlObj1 += '<thead style ="background: #CFCFCF; background: -moz-linear-gradient(top, #dbdbdb 0%, #d3d3d3 66%, #CFCFCF 100%); background: -webkit-linear-gradient(top, #dbdbdb 0%, #d3d3d3 66%, #CFCFCF 100%);  background: linear-gradient(to bottom, #dbdbdb 0%, #d3d3d3 66%, #CFCFCF 100%);  border-bottom: 1px solid #989898;">';
			htmlObj1 += '<tr>';
			htmlObj1 += '<th colspan="6" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>'+month_Name.substring(0, 3)+' Fee Reconciliation -'+customer_name+'  </b></center></th>';
			//htmlObj1 +='</tr>';
			htmlObj1 += '<th colspan="3" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Project SOW Estimate</b></center></th>';
			//htmlObj1 +='</tr>';
			htmlObj1 += '<th colspan="3" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  '+yearId+' Brief Estimate </b></center></th>';
			htmlObj1 += '<th colspan="1" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Variance </b></center></th>';
			htmlObj1 += '<th colspan="3" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  Project Actuals  </b></center></th>';
			
			htmlObj1 += '<th colspan="1" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  Variance   </b></center></th>';
			htmlObj1 += '<th colspan="1" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  </b></center></th>';
			htmlObj1 += '<th colspan="1" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  YTD Billing(fees and OOP) </b></center></th>';
			htmlObj1 += '<th colspan="12" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  Actual Fees and Forecast 	</b></center></th>';
			htmlObj1 += '<th colspan="12" align="center" style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Actual OOPs and Forecast</b></center></th>';
			htmlObj1 += '</tr>';

											
			htmlObj1 += '<tr>';
			
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Client PO Number</center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Project Number</center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Client : Project</center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Status</center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Start Date</center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Projected End Date</center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Fee $</b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> OOP $ </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Total Fee and OOP $ </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Estimated Fees </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Estimated OOP </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Estimate Total  </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> (Over)/Under </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  Actual Fee  </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  Actual OOP </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  Actual Total  </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> (Over)/Under  </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>  % Utilized  </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Billed to Date </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Jan </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Feb </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Mar </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Apr </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> May </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Jun </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Jul </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Aug </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Sep </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Oct </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Nov </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Dec </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Jan </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Feb </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Mar </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Apr </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> May </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Jun </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b> Jul </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Aug </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Sep </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Oct </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Nov </b></center></th>';
			htmlObj1 += '<th style="border: 1px solid #000000; padding: 5px 4px;"><center><b>Dec </b></center></th>';
		
		htmlObj1 += '</tr>';
		htmlObj1 += '</thead>';
		htmlObj1 += '<tbody>';
		var searchCount=5;
		 var cliNo= "US-11066-1";
		 var clinetID =2091;
		if(gstIn && yearId && gstCustomerId)
		{
			var mum_value = gstIn.split(",");
				log.debug('mum_value.length',mum_value.length)
				var year_id = yearId.split(",")
				for(var jk=0;jk<year_id.length;jk++)
				{
					for(var i=0;i<mum_value.length;i++)
					{
						log.debug("mum_value[i]",mum_value[i])
						var id = serchDataByClient(mum_value[i],gstCustomerId,year_id[jk]);
					}
				}
			/*	for(var i=0;i<jsonAyyaData.length;i++)
				{
					log.debug('jsonAyyaData[i].cli_po_num============',jsonAyyaData[i].cli_po_num)
				}
				*/
				log.debug('jsonAyyaData.length',jsonAyyaData.length)
			var htmlStr1 = "";
			var total_Sow_Opp = 0;
			var customrecord_sss_proj_snapshot = [];
			var project_Number_Array = [];
			try{
			for(var x=0;x<jsonAyyaData.length;x++)
			 {
				 log.debug('x',x)
				 log.debug('jsonAyyaData',jsonAyyaData[x])
					//	r1c1 = result.getValue(result.columns[0]);//
					var Client_PO_Number = jsonAyyaData[x].Client_PO_Number;//result.getValue(result.columns[1]); //Client PO Number
					log.debug('Client_PO_Number',Client_PO_Number)
					var Project_Number =jsonAyyaData[x].Project_Number; //result.getValue(result.columns[3]); //Project Number
					log.debug('Project_Number',Project_Number)
					project_Number_Array.push(Project_Number);
					var Client_Project = jsonAyyaData[x].Client_Project;//result.getValue(result.columns[2]); //Client : Project
					var Status = jsonAyyaData[x].Status;//result.getText({name: "entitystatus",join: "CUSTRECORD_SSS_SS_PROJ",label: "Status"});
					var Start_Date = jsonAyyaData[x].Start_Date;//result.getValue(result.columns[4]);
					var Projected_End_Date = jsonAyyaData[x].Projected_End_Date;//result.getValue(result.columns[5]);

					var sowFee = jsonAyyaData[x].sowFee;//result.getValue(result.columns[8]);
					if(sowFee) {
						sowFee = sowFee;
					}
					else {
						sowFee = 0
					}
					var sowOpp = jsonAyyaData[x].sowOpp;//result.getValue(result.columns[9]);
					if(sowOpp) {
						sowOpp = sowOpp;
					}
					else {
						sowOpp = 0;
					}
					total_Sow_Opp = Number(sowOpp) + Number(sowFee)


					var ACTUAL_FEE =jsonAyyaData[x].ACTUAL_FEE;// result.getValue(result.columns[13]);
					var ACTUAL_OOP = jsonAyyaData[x].ACTUAL_OOP;//result.getValue(result.columns[14]);
					var actual_TOTAL = jsonAyyaData[x].actual_TOTAL;//result.getValue(result.columns[15]);
					var ESTIMATED_FEES =jsonAyyaData[x].ESTIMATED_FEES;// result.getValue(result.columns[10]);
					var ESTIMATED_OOP =jsonAyyaData[x].ESTIMATED_OOP;// result.getValue(result.columns[11]);
					var ESTIMATE_TOTAL =jsonAyyaData[x].ESTIMATE_TOTAL;// result.getValue(result.columns[12]);
					//log.debug('actual_TOTAL',actual_TOTAL)
				//	log.debug('ESTIMATE_TOTAL',ESTIMATE_TOTAL)
					
					var Over_Under = Math.abs(Number(ESTIMATE_TOTAL)-Number(total_Sow_Opp))
					var over_under = Math.abs(Number(ESTIMATE_TOTAL)-Number(actual_TOTAL))
					var utilized = jsonAyyaData[x].utilized;
					log.debug('utilized',utilized)
					if(_logValidation(utilized))
					{
						utilized = utilized;
						
					}else{
						utilized=0;
					}
					/*if(Number(actual_TOTAL) >0 && Number(ESTIMATE_TOTAL) >0)
					{
						var newValue =Number(actual_TOTAL)/Number(ESTIMATE_TOTAL)
						var utilized = newValue*100;
						log.debug('utilized',utilized)
					}*/
					var Billed_to_Date = jsonAyyaData[x].Billed_to_Date;//result.getValue(result.columns[16]);
					var cli_po_num =jsonAyyaData[x].cli_po_num; /*result.getValue({
						name: "custentity_mp_client_po_num",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Client PO Number"
					});*/
					log.debug('cli_po_num', cli_po_num)
					var year_Value =jsonAyyaData[x].yearValue;
					/*********second Search*************/

					// for second serach data
					var monthlyArray = secordSearch(gstCustomerId,cli_po_num,Project_Number,year_Value);
					log.debug('monthlyArray for result',monthlyArray)
					//log.debug('newAr', newAr)

					result = [];

					monthlyArray.forEach(function(a) {
						if(!this[a.projectID]) {
							this[a.projectID] = {
								projectID: a.projectID,
								itemType: '',
								January: '',
								February: '',
								March: '',
								April: '',
								May: '',
								June: '',
								July: '',
								August: '',
								September: '',
								October: '',
								November: '',
								December: ''
							};
							result.push(this[a.projectID]);
						}
						this[a.projectID].itemType += a.itemType + '*';
						this[a.projectID].January += a.January + '*';
						this[a.projectID].February += a.February + '*';
						this[a.projectID].March += a.March + '*';
						this[a.projectID].April += a.April + '*';
						this[a.projectID].May += a.May + '*';
						this[a.projectID].June += a.June + '*';
						this[a.projectID].July += a.July + '*';
						this[a.projectID].August += a.August + '*';
						this[a.projectID].September += a.September + '*';
						this[a.projectID].October += a.October + '*';
						this[a.projectID].November += a.November + '*';
						this[a.projectID].December += a.December + '*';

					}, Object.create(null));
					log.debug('result', result)

					var keys = []
					var values = []
					i = 0
					for(var key in result[0]) {
						keys[i] = key;
						values[i] = result[0][key]
						i = i + 1
					}
					log.debug('keys', keys);
					log.debug('values', values)
					
					
					
					
					var crrMonthTotal=0;
					var crrOppMonthTotal=0;
					
					var chekJune=0;
					var rJanuary = 0;
					var rFebruary = 0;
					var rMarch = 0;
					var rApril = 0;
					var rMay = 0;
					var rJune = 0;
					var rJuly = 0;
					var rAugust = 0;
					var rSeptember = 0;
					var rOctober = 0;
					var rNovember = 0;
					var rDecember = 0;

					var oppJanuary = 0;
					var oppFebruary = 0;
					var oppMarch = 0;
					var oppApril =0;
					var oppMay = 0;
					var oppJune = 0;
					var oppJuly = 0;
					var oppAugust = 0;
					var oppSeptember =0;
					var oppOctober = 0;
					var oppNovember = 0;
					var oppDecember = 0;

					for(var i = 0; i < keys.length; i++) {
						if(keys[i] != monthName) {
							if(i == 0 || i == 1) {

							}
							else {
								//log.debug('keys-', keys[i]);
								var i_Item = values[1];
								//	log.debug('i_Item',i_Item);
								//var r_Item = i_Item.substring(0, i_Item.length - 1);
								var split_Item = i_Item.split("*");
								//log.debug('item length',split_Item);
								// for fees

								var fr1c1 = values[2];

								var fr1_c1 = fr1c1.split("*");
								//	log.debug('fr1_c1',fr1_c1)

								var fr1c2 = values[3];
								//log.debug('fr1c2',fr1c2)
								var fr1_c2 = fr1c2.split("*");

								var fr1c3 = values[4];
								var fr1_c3 = fr1c3.split("*");

								var fr1c4 = values[5];
								var fr1_c4 = fr1c4.split("*");

								var fr1c5 = values[6];
								var fr1_c5 = fr1c5.split("*");

								var fr1c6 = values[7];
								var fr1_c6 = fr1c6.split("*");

								var fr1c7 = values[8];
								var fr1_c7 = fr1c7.split("*");

								var fr1c8 = values[9];
								var fr1_c8 = fr1c8.split("*");

								var fr1c9 = values[10];
								var fr1_c9 = fr1c9.split("*");

								var fr1c10 = values[11];
								var fr1_c10 = fr1c10.split("*");

								var fr1c11 = values[12];
								var fr1_c11 = fr1c11.split("*");

								var fr1c12 = values[13];
								var fr1_c12 = fr1c12.split("*");

								for(var j = 0; j < split_Item.length; j++) {
									if(split_Item[j] == "Fees" ) {
										if(keys[i] === 'January' && Ja<=m_Date ) {

											rJanuary = fr1_c1[j];
										//	log.debug('rJanuary', rJanuary)
									//	crrMonthTotal+=Number(rJanuary)
										
										}
										if(keys[i] === 'February' && Fe<=m_Date) {
											rFebruary = fr1_c2[j];
										//	log.debug('rFebruary', rFebruary)
										crrMonthTotal+=Number(rFebruary)
										
										}
										if(keys[i] === 'March' && Mr<=m_Date) {
											rMarch = fr1_c3[j];
										//	log.debug('rMarch', rMarch)
										crrMonthTotal+=Number(rMarch)
										
										}
										if(keys[i] === 'April' && Ap<=m_Date) {
											rApril = fr1_c4[j];
										//	log.debug('rApril', rApril)
										crrMonthTotal+=Number(rApril)
										
										}
										if(keys[i] === 'May' && Ma<=m_Date) {
											rMay = fr1_c5[j];
											//oppMay =fr1_c5[j+1]
											//log.debug('in fees',oppMay)
											//log.debug('rMay', rMay)
											crrMonthTotal+=Number(rMay)
											
										}
										if(keys[i] === 'June' && Ju<=m_Date) {
											rJune = fr1_c6[j];
											crrMonthTotal+=Number(rJune)
											
										}
										if(keys[i] === 'July' && Jl<=m_Date) {
											rJuly = fr1_c7[j];
											crrMonthTotal+=Number(rJuly)
											
										//	log.debug('rJuly', rJuly)
										}
										if(keys[i] === 'August' && Au<=m_Date) {
											rAugust = fr1_c8[j];
											crrMonthTotal+=Number(rAugust)
											
										//	log.debug('rAugust', rAugust)
										}
										if(keys[i] === 'September' && Se<=m_Date) {
											rSeptember = fr1_c9[j];
											crrMonthTotal+=Number(rSeptember)
											
										//	log.debug('rSeptember', rSeptember)
										}
										if(keys[i] === 'October' && Oc<=m_Date) {
											rOctober = fr1_c10[j];
											crrMonthTotal+=Number(rOctober)
											
										//	log.debug('rOctober', rOctober)
										}
										if(keys[i] === 'November' && No<=m_Date) {
											rNovember = fr1_c11[j];
											crrMonthTotal+=Number(rNovember)
											
										//	log.debug('rNovember', rNovember)
										}
										if(keys[i] === 'December' && De<=m_Date) {
											rDecember = fr1_c12[j];
											crrMonthTotal+=Number(rDecember)
											
										//	log.debug('rDecember', rDecember)
										}
									} // for opp
									else if(split_Item[j]=="Out of Pockets (OOP)"){
										if(keys[i] === 'January' && Ja<=m_Date) 
										{

											oppJanuary = fr1_c1[j];
											log.debug('oppJanuary', oppJanuary)
										}
										if(keys[i] === 'February' && Fe<=m_Date) {
											oppFebruary = fr1_c2[j];
											log.debug('oppFebruary', oppFebruary)
											crrOppMonthTotal+=Number(oppFebruary)
										}
										if(keys[i] === 'March' && Mr<=m_Date) {
											oppMarch = fr1_c3[j];
											log.debug('oppMarch', oppMarch)
											crrOppMonthTotal+=Number(oppMarch)
										}
										if(keys[i] === 'April' && Ap<=m_Date) {
											oppMay = fr1_c4[j];
											log.debug('oppApril', oppApril)
											crrOppMonthTotal+=Number(oppMay)
										}
										if(keys[i] === 'May' && Ma<=m_Date) {
											oppMay = fr1_c5[j];
											log.debug('oppMay', oppMay)
											crrOppMonthTotal+=Number(oppMay)
										}
										if(keys[i] === 'June' && Ju<=m_Date) {
											oppJune = fr1_c6[j];
											log.debug('oppJune', oppJune)
											crrOppMonthTotal+=Number(oppJune)
										}
										if(keys[i] === 'July' && Jl<=m_Date) {
											oppJuly = fr1_c7[j];
											log.debug('oppJuly', oppJuly)
											crrOppMonthTotal+=Number(oppJuly)
										}
										if(keys[i] === 'August' && Au<=m_Date) {
											oppAugust = fr1_c8[j];
											log.debug('oppAugust', oppAugust)
											crrOppMonthTotal+=Number(oppAugust)
										}
										if(keys[i] === 'September' && Se<=m_Date) {
											oppSeptember = fr1_c9[j];
											log.debug('oppSeptember', oppSeptember)
											crrOppMonthTotal+=Number(oppSeptember)
										}
										if(keys[i] === 'October' && Oc<=m_Date) {
											oppOctober = fr1_c10[j];
											log.debug('oppOctober', oppOctober)
											crrOppMonthTotal+=Number(oppOctober)
										}
										if(keys[i] === 'November' && No<=m_Date) {
											oppNovember = fr1_c11[j];
											log.debug('oppNovember', oppNovember)
											crrOppMonthTotal+=Number(oppNovember)
										}
										if(keys[i] === 'December' && De<=m_Date) {
											oppDecember = fr1_c12[j];
											log.debug('oppDecember', oppDecember)
											crrOppMonthTotal+=Number(oppDecember)
										}
									}
								//	log.debug('crrMonthTotal',crrMonthTotal)
								}
							}
						}
						else {
							break;
						}
					}
					
					/************** Project Forecast Fee and OOP: Results *************/

				var oppMonthArray=thirdSerchForOpp(gstCustomerId,cli_po_num,Project_Number,year_Value);
					log.debug('result data oppMonthArray',oppMonthArray)
					opp_result = [];

					oppMonthArray.forEach(function(a) {
						if(!this[a.projectID]) {
							this[a.projectID] = {
								projectID: a.projectID,
								itemType: '',
								January: '',
								February: '',
								March: '',
								April: '',
								May: '',
								June: '',
								July: '',
								August: '',
								September: '',
								October: '',
								November: '',
								December: ''
							};
							opp_result.push(this[a.projectID]);
						}
						this[a.projectID].itemType += a.itemType + '*';
						this[a.projectID].January += a.January + '*';
						this[a.projectID].February += a.February + '*';
						this[a.projectID].March += a.March + '*';
						this[a.projectID].April += a.April + '*';
						this[a.projectID].May += a.May + '*';
						this[a.projectID].June += a.June + '*';
						this[a.projectID].July += a.July + '*';
						this[a.projectID].August += a.August + '*';
						this[a.projectID].September += a.September + '*';
						this[a.projectID].October += a.October + '*';
						this[a.projectID].November += a.November + '*';
						this[a.projectID].December += a.December + '*';

					}, Object.create(null));
					log.debug('opp_result', opp_result);
					var keysOpp = []
					var valuesOpp = []
					i = 0
					for(var key in opp_result[0]) {
						keysOpp[i] = key;
						valuesOpp[i] = opp_result[0][key]
						i = i + 1
					}
					log.debug('keysOpp', keysOpp);
					log.debug('valuesOpp', valuesOpp);


					var checkDate = new Date();
					var d_date =checkDate.getMonth()+1;
					
					for(var i = 0; i < keysOpp.length; i++) {
						if(i < (d_date+2)) {
							//log.debug('keysOpp[i]',keysOpp[i])
							if(i == 0 || i == 1) {

							}
						}
						else {
						//	log.debug('keysOpp-', keysOpp[i]);
							var i_Item = valuesOpp[1]
							//var r_Item = i_Item.substring(0, i_Item.length - 1);
							var split_Item = i_Item.split("*")
							//	log.debug('item length',split_Item.length);
							//log.debug('split_Item[0]',split_Item[0]);

							var r1c1 = valuesOpp[2];
							var r1_c1 = r1c1.split("*");

							var r1c2 = valuesOpp[3];
							var r1_c2 = r1c2.split("*");

							var r1c3 = valuesOpp[4];
							var r1_c3 = r1c3.split("*");

							var r1c4 = valuesOpp[5];
							var r1_c4 = r1c4.split("*");

							var r1c5 = valuesOpp[6];
							var r1_c5 = r1c5.split("*");

							var r1c6 = valuesOpp[7];
							var r1_c6 = r1c6.split("*");

							var r1c7 = valuesOpp[8];
							var r1_c7 = r1c7.split("*");

							var r1c8 = valuesOpp[9];
							var r1_c8 = r1c8.split("*");

							var r1c9 = valuesOpp[10];
							var r1_c9 = r1c9.split("*");

							var r1c10 = valuesOpp[11];
							var r1_c10 = r1c10.split("*");

							var r1c11 = valuesOpp[12];
							var r1_c11 = r1c11.split("*");

							var r1c12 = valuesOpp[13];
							var r1_c12 = r1c12.split("*");



							for(var s = 0; s < split_Item.length; s++) {
								if(split_Item[s] == '1') {
									if(keysOpp[i] === 'January' && Ja>=m_Date) {
										rJanuary = r1_c1[s];
									//crrOppMonthTotal+=Number(rJanuary)
									//	log.debug('rJanuary', rJanuary)
									}
									if(keysOpp[i] === 'February' && Fe>=m_Date) {

										rFebruary = r1_c2[s];
									//	crrOppMonthTotal+=Number(rFebruary)
									//	log.debug('rFebruary', rFebruary)
									}
									if(keysOpp[i] === 'March' && Mr>=m_Date) {


										rMarch = r1_c3[s];
										//crrOppMonthTotal+=Number(rMarch)
									//	log.debug('rMarch', rMarch)
									}
									if(keysOpp[i] === 'April' && Ap>=m_Date) {

										rApril = r1_c4[s];
										crrOppMonthTotal+=Number(rApril)
									//	log.debug('rApril', rApril)
									}
									if(keysOpp[i] === 'May' && Ma>=m_Date) {

										rMay = r1_c5[s];
										//crrOppMonthTotal+=Number(rMay)
										//rMay = valuesOpp[i];
									//	log.debug('rMay', rMay)
									}
									if(keysOpp[i] === 'June' && Ju>=m_Date) {

										rJune = r1_c6[s];
										rJune = valuesOpp[i];
										//crrOppMonthTotal+=Number(rJune)
									//	log.debug('rJune', rJune)
									}
									if(keysOpp[i] === 'July' && Jl>=m_Date) {

										rJuly = r1_c7[s];
									//	crrOppMonthTotal+=Number(rJuly)
										//rJuly = valuesOpp[i];
									//	log.debug('rJuly', rJuly)
									}
									if(keysOpp[i] === 'August' && Au>=m_Date) {

										rAugust = r1_c8[s];
									//	crrOppMonthTotal+=Number(rAugust)
										//rAugust = valuesOpp[i];
									//	log.debug('rAugust', rAugust)
									}
									if(keysOpp[i] === 'September' && Se>=m_Date) {

										rSeptember = r1_c9[s];
										//crrOppMonthTotal+=Number(rJanuary)
										//rSeptember = valuesOpp[i];
									//	log.debug('rSeptember', rSeptember)
									}
									if(keysOpp[i] === 'October' && Oc>=m_Date) {

										rOctober = r1_c10[s];
									//	crrOppMonthTotal+=Number(rOctober)
										//rOctober = valuesOpp[i];
									//	log.debug('rOctober', rOctober)
									}
									if(keysOpp[i] === 'November' && No>=m_Date) {

										rNovember = r1_c11[s];
									//	crrOppMonthTotal+=Number(rNovember)
										//rNovember = valuesOpp[i];
									//	log.debug('rNovember', rNovember)
									}
									if(keysOpp[i] === 'December' && De>=m_Date) {

										rDecember = r1_c12[s];
									//	crrOppMonthTotal+=Number(rDecember)

										//rDecember = valuesOpp[i];
									//	log.debug('rDecember', rDecember)
									}
								}
								else if(split_Item[s] == '2') {
									//log.debug('for opp')
									if(keysOpp[i] === 'January' && Ja>=m_Date) {

										oppJanuary = r1_c1[s];
									//	log.debug('oppJanuary', oppJanuary)
									}
									if(keysOpp[i] === 'February'&& Fe>=m_Date) {
										oppFebruary = r1_c2[s];
										//crrOppMonthTotal+=Number(oppFebruary)
										//log.debug('oppFebruary', oppFebruary)
									}
									if(keysOpp[i] === 'March' && Mr>=m_Date) {
										oppMarch = r1_c3[s];
									//	log.debug('oppMarch', oppMarch)
									//crrOppMonthTotal+=Number(oppMarch)
									}
									if(keysOpp[i] === 'April' && Ap>=m_Date) {
										oppApril = r1_c4[s];;
									//	log.debug('oppApril', oppApril)
									//crrOppMonthTotal+=Number(oppApril)
									}
									if(keysOpp[i] === 'May' && Ma>=m_Date) {
										oppMay = r1_c5[s];
									//	log.debug('oppMay', oppMay)
									//crrOppMonthTotal+=Number(oppMay)
									}
									if(keysOpp[i] === 'June' && Ju>=m_Date) {
										oppJune = r1_c6[s];
									//	log.debug('oppJune', oppJune)
									//crrOppMonthTotal+=Number(oppJune)
									}
									if(keysOpp[i] === 'July' && Jl>=m_Date) {
										oppJuly = r1_c7[s];
									//	log.debug('oppJuly', oppJuly)
									//crrOppMonthTotal+=Number(oppJuly)
									}
									if(keysOpp[i] === 'August' && Au>=m_Date) {
										oppAugust = r1_c8[s];
									//	log.debug('oppAugust', oppAugust)
									//crrOppMonthTotal+=Number(oppAugust)
									}
									if(keysOpp[i] === 'September' && Se>=m_Date) {
										oppSeptember = r1_c9[s];
									//	log.debug('oppSeptember', oppSeptember)
									//crrOppMonthTotal+=Number(oppSeptember)
									}
									if(keysOpp[i] === 'October' && Oc>=m_Date) {
										oppOctober = r1_c10[s];
									//	log.debug('oppOctober', oppOctober)
									//crrOppMonthTotal+=Number(oppOctober)
									}
									if(keysOpp[i] === 'November' && No>=m_Date) {
										oppNovember = r1_c11[s];
									//	log.debug('oppNovember', oppNovember)
									//crrOppMonthTotal+=Number(oppNovember)
									}
									if(keysOpp[i] === 'December' && De>=m_Date) {
										oppDecember = r1_c12[s];
									//	log.debug('oppDecember', oppDecember)
									//crrOppMonthTotal+=Number(oppDecember)
									}

								}


							}
						}
					}
				var c_p_name = Client_Project.split(":");
				var c_id = '';
				for(var p=1;p<c_p_name.length;p++)
				{
					c_id+=c_p_name[p]+" ";
				}
					/********************** secnond search****************/

					htmlObj1 += '<tr>';
					htmlObj1 += '<td style="border: 1px solid #000000;">' + cli_po_num + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">' + Project_Number + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">' + c_id + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">' + Status + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">' + Start_Date + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">' + Projected_End_Date + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( sowFee).toFixed() + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(sowOpp).toFixed()+ '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(total_Sow_Opp).toFixed() + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( ESTIMATED_FEES).toFixed()+ '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(ESTIMATED_OOP).toFixed() + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(ESTIMATE_TOTAL).toFixed()+ '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(Over_Under).toFixed() + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(ACTUAL_FEE).toFixed()+ '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(ACTUAL_OOP).toFixed() + '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(actual_TOTAL).toFixed()+ '</td>';
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(over_under).toFixed()+ '</td>';
					if(utilized)
					{
						htmlObj1 += '<td style="border: 1px solid #000000;">' + utilized.toFixed() + '%</td>';
					}else{
						htmlObj1 += '<td style="border: 1px solid #000000;">0%</td>';
					}
					var newValueTotal = Number(ACTUAL_FEE).toFixed()-Number(crrMonthTotal)
					var oppValueTotal = Number(ACTUAL_OOP).toFixed() - Number(crrOppMonthTotal)
					log.debug('oppValueTotal',oppValueTotal)
					
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(Billed_to_Date).toFixed() + '</td>';
					//htmlObj1 += '<td style="border: 1px solid #000000;"></td>';
					//if(retrun_id)
					{
						
						if(Fe==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
					else if(rFebruary)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rFebruary).toFixed() + '</td>';
							//subTotalCrrMonth+=newValueTotal
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Mr==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rMarch)
						{
								htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rMarch).toFixed() + '</td>';
							//	subTotalCrrMonth+=newValueTotal
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Ap==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rApril)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rApril).toFixed() + '</td>';
							//subTotalCrrMonth+=newValueTotal
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Ma==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
					else if(rMay)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number(rMay).toFixed() + '</td>';
							//subTotalCrrMonth+=newValueTotal
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Ju==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rJune)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rJune).toFixed() + '</td>';
							//subTotalCrrMonth+=newValueTotal
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Jl==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rJuly)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rJuly).toFixed() + '</td>';
							
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Au ==(Number(m_Date)-1))
							{
								htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
								subTotalCrrMonth+=newValueTotal
								
							}
						else if(rAugust)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rAugust).toFixed() + '</td>';
							
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Se==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rSeptember)
						{
								htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rSeptember).toFixed() + '</td>';
								
						}
						else{
								htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
					if(Oc==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rOctober)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rOctober).toFixed()+ '</td>';
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(No==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
					else if(rNovember)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(rNovember).toFixed()+ '</td>';
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(De==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(newValueTotal).toFixed() + '</td>';
							subTotalCrrMonth+=newValueTotal
						}
						else if(rDecember)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( rDecember).toFixed() + '</td>';
						}
						else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						// for opp
					//	htmlObj1 += '<td style="border: 1px solid #000000;">' + oppJanuary + '</td>';
					
					// for forecoast result
					if(Fe==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number(oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal;
						}
						else if(oppFebruary)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppFebruary).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Mr==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else	if(oppMarch)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppMarch).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
							if(Ap==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else if(oppApril)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppApril).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Ma==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else if(oppMay)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppMay).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
							if(Ju==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else if(oppJune)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppJune).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
							if(Jl==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else if(oppJuly)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppJuly).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
							if(Au==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else if(oppAugust)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppAugust).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Se==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else	if(oppSeptember)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppSeptember).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(Oc==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else	if(oppOctober)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppOctober).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(No==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else	if(oppNovember)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppNovember).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
						if(De==(Number(m_Date)-1))
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' +Number( oppValueTotal).toFixed() + '</td>';
							subTotalOppMonth+=oppValueTotal
						}
						else	if(oppDecember)
						{
							htmlObj1 += '<td style="border: 1px solid #000000;">$' + Number(oppDecember).toFixed() + '</td>';
						}else{
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';
						}
							
							htmlObj1 += '<td style="border: 1px solid #000000;">$0</td>';

						}
						

					htmlObj1 += '</tr>';
	//	totalValue+=Over_Under;
		// for additon of all tha maount
					 c1+=Number(sowFee);
					 c2+=Number(sowOpp);
					 c3+=Number(total_Sow_Opp);
					 c4+=Number(ESTIMATED_FEES);
					 c5+=Number(ESTIMATED_OOP);
					 c6+=Number(ESTIMATE_TOTAL);
					 c7+=Number(Over_Under);
					 c8+=Number(ACTUAL_FEE);
					 c9+=Number(ACTUAL_OOP);
					 c10+=Number(actual_TOTAL);
					 c11+=Number(over_under);
					 if(utilized)
					 {
						  c12+=Number(Number(utilized).toFixed());
					 }
					
					 c13+=Number(Billed_to_Date);
					 c14+=Number(rFebruary );
					 c15+=Number(rMarch    );
					 c16+=Number(rApril    );
					 c17+=Number(rMay      );
					 c18+=Number(rJune     );
					 c19+=Number(rJuly     );
					 c20+=Number(rAugust   );
					 c21+=Number(rSeptember);
					 c22+=Number(rOctober  );
					 c23+=Number(rNovember );
					 c24+=Number(rDecember );
					 c38+=0;
					 c25+=Number(oppFebruary );
					 c26+=Number(oppMarch    );
					 c27+=Number(oppApril    );
					 c28+=Number(oppMay      );
					 c29+=Number(oppJune     );
					 c30+=Number(oppJuly     );
					 c31+=Number(oppAugust   );
					 c32+=Number(oppSeptember);
					 c33+=Number(oppOctober  );
					 c34+=Number(oppNovember );
					 c36+=Number(oppDecember );
					 c37+=0;
		
		
		var empty='';
					excelObj  += '<Row>' ;
					excelObj+=	'<Cell><Data ss:Type="String">' + cli_po_num + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">' + Project_Number + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">' + c_id + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">' + Status + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">' + Start_Date + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">' + Projected_End_Date + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(sowFee).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(sowOpp).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(total_Sow_Opp ).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(ESTIMATED_FEES ).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(ESTIMATED_OOP ).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(ESTIMATE_TOTAL).toFixed() + '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(Over_Under).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(ACTUAL_FEE).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(ACTUAL_OOP).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(actual_TOTAL ).toFixed()+ '</Data></Cell>' ;
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(over_under).toFixed() + '</Data></Cell>' ;
					if(utilized)
					{
						excelObj+=	'<Cell><Data ss:Type="String">' +  Number(utilized).toFixed()+ '%</Data></Cell>' ;
					}else{
						excelObj+=	'<Cell><Data ss:Type="String">0%</Data></Cell>' ;
					}
					
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(Billed_to_Date ).toFixed()+ '</Data></Cell>' ;
						/******addon**********/
						
							
					if(Fe==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rFebruary).toFixed() + '</Data></Cell>' ;
						}
								
					if(Mr==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rMarch ).toFixed()+ '</Data></Cell>' ;
						}
						if(Ap==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rApril).toFixed() + '</Data></Cell>' ;
						}
						if(Ma==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' +Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rMay ).toFixed()+ '</Data></Cell>' ;
						}
						
						if(Ju==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rJune ).toFixed()+ '</Data></Cell>' ;
						}
						
						if(Jl==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
								excelObj+='<Cell><Data ss:Type="String">$' + Number(rJuly ).toFixed()+ '</Data></Cell>' ;
						}
						if(Au==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rAugust ).toFixed()+ '</Data></Cell>' ;
						}
						if(Se==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rSeptember).toFixed() + '</Data></Cell>' ;
						}
						if(Oc==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rOctober).toFixed() + '</Data></Cell>' ;
						}
						if(No==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rNovember).toFixed()+ '</Data></Cell>' ;
						}
						if(De==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(newValueTotal).toFixed()+ '</Data></Cell>' ;
						}else{
							excelObj+='<Cell><Data ss:Type="String">$' + Number(rDecember).toFixed()+ '</Data></Cell>' ;
						}
					excelObj+='<Cell><Data ss:Type="String">$0</Data></Cell>' ;
						
						
						
					// for opp
							
					if(Fe==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
							excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppFebruary).toFixed()+ '</Data></Cell>' ;
						}
					if(Ma==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
							excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppMarch).toFixed()+ '</Data></Cell>' ;
						}
					if(Ap==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppApril).toFixed() + '</Data></Cell>' ;
						}
						if(Ma==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppMay).toFixed() + '</Data></Cell>' ;
						}
						if(Ju==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppJune).toFixed() + '</Data></Cell>' ;
						}
						if(Jl==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppJuly).toFixed() + '</Data></Cell>' ;
						}
						if(Au==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppAugust).toFixed() + '</Data></Cell>' ;
						}
						if(Se==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppSeptember).toFixed() + '</Data></Cell>' ;
						}
						if(Oc==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppOctober ).toFixed()+ '</Data></Cell>' ;
						}
						if(No==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppNovember).toFixed() + '</Data></Cell>' ;
						}
						if(De==(Number(m_Date)-1))
						{
					excelObj+='<Cell><Data ss:Type="String">$' + Number(oppValueTotal).toFixed() + '</Data></Cell>' ;
						}else{
					excelObj+=	'<Cell><Data ss:Type="String">$' + Number(oppDecember).toFixed() + '</Data></Cell>' ;
						}
					excelObj+=	'<Cell><Data ss:Type="String">$0</Data></Cell>' ;
					excelObj+=	'</Row>';
				
					continue;
				}
			}catch(e)
			{
				log.error("e.messsage",e.message)
			}
			var utTotal = c10 /c6
			var urPercentage = utTotal*100
			if(_logValidation(urPercentage))
			{
				urPercentage=urPercentage;
			}else{
				urPercentage=0
			}
				htmlObj1 += '<tr>';
				htmlObj1 += '<td colspan="6" style="border: 1px solid #000000;">Total</td>';
				/*htmlObj1 += '<td style="border: 1px solid #000000;"></td>';
				htmlObj1 += '<td style="border: 1px solid #000000;"></td>';
				htmlObj1 += '<td style="border: 1px solid #000000;"></td>';
				htmlObj1 += '<td style="border: 1px solid #000000;"></td>';
				htmlObj1 += '<td style="border: 1px solid #000000;"></td>';
				*/
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c1.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c2.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c3.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c4.toFixed() + '</td>';                          
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c5.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' +c6.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c7.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c8.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c9.toFixed() + '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c10.toFixed()+ '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c11.toFixed()+ '</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">' + urPercentage.toFixed()+ '%</td>';
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c13.toFixed()+ '</td>';
				
				if(Ja ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c14.toFixed()+ '</td>';
				}
				if(Fe ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}
				else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c15.toFixed()+ '</td>';
				}
				if(Mr ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c16.toFixed()+ '</td>';
				}
				if(Ap ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c17.toFixed()+ '</td>';
				}
				if(Ma ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c18.toFixed()+ '</td>';
				}
				if(Ju ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else
				{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c19.toFixed()+ '</td>';
				}
				if(Jl ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c20.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c20.toFixed()+ '</td>';
				}
				if(Au ==(Number(m_Date)-2))
				{
					//subTotalCrrMonth
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 +='<td style="border: 1px solid #000000;">$' + c21.toFixed()+ '</td>';
				}
				if(Se ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else
				{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c22.toFixed()+ '</td>';
				}
				if(Oc ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c23.toFixed()+ '</td>';
				}
				if(No ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c24.toFixed()+ '</td>';
				}
				if(De ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalCrrMonth.toFixed()+ '</td>';
				}else{
					htmlObj1 += '<td style="border: 1px solid #000000;">$' + c38.toFixed()+ '</td>';
				}
				
				// for opp
				if(Ja ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c25.toFixed()+ '</td>';
				}
				if(Fe ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c26.toFixed()+ '</td>';
				}
				if(Mr ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c27.toFixed()+ '</td>';
				}
				if(Ap ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c28.toFixed()+ '</td>';
				}
				if(Ma ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c29.toFixed()+ '</td>';
				}
				if(Ju ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c30.toFixed()+ '</td>';
				}
				if(Jl ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c31.toFixed()+ '</td>';
				}
				if(Au ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c32.toFixed()+ '</td>';
				}
				if(Se ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c33.toFixed()+ '</td>';
				}
				if(Oc ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c34.toFixed()+ '</td>';
				}
				if(No ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c36.toFixed()+ '</td>';
				}
				if(De ==(Number(m_Date)-2))
				{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + subTotalOppMonth.toFixed()+ '</td>';
				}else{
				htmlObj1 += '<td style="border: 1px solid #000000;">$' + c37.toFixed()+ '</td>';
				}
			
				htmlObj1 += '</tr>';
				
				
				// for show in sheet
					excelObj+= '<Row>' ;
					excelObj  +='<Cell ss:MergeAcross="5"><Data ss:Type="String">Total</Data></Cell>' ;
						/*'<Cell><Data ss:Type="String"></Data></Cell>' +
						'<Cell><Data ss:Type="String"></Data></Cell>' +
						'<Cell><Data ss:Type="String"></Data></Cell>' +
						'<Cell><Data ss:Type="String"></Data></Cell>' +
						'<Cell><Data ss:Type="String">Total</Data></Cell>' +*/
					excelObj  +='<Cell><Data ss:Type="String">$' +  c1.toFixed()+ '</Data></Cell>' ;
					excelObj  +='<Cell><Data ss:Type="String">$' +  c2.toFixed()+ '</Data></Cell>' ;
					excelObj  +='<Cell><Data ss:Type="String">$' +  c3.toFixed() + '</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">$' +  c4.toFixed() + '</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">$' +  c5.toFixed()+ '</Data></Cell>' ;
					excelObj  +='<Cell><Data ss:Type="String">$' +  c6.toFixed() + '</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">$' +  c7.toFixed() + '</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">$' + c8.toFixed() + '</Data></Cell>' ;
					excelObj  +='<Cell><Data ss:Type="String">$' + c9.toFixed() + '</Data></Cell>' ;
					excelObj  +='<Cell><Data ss:Type="String">$' + c10.toFixed() + '</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">$' + c11.toFixed() + '</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">' + urPercentage.toFixed() + '%</Data></Cell>';
					excelObj  +='<Cell><Data ss:Type="String">$' + c13.toFixed() + '</Data></Cell>';
					/******addon**********/
						
						
						
						if(Ja ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
						excelObj  +=	'<Cell><Data ss:Type="String">$' + c14.toFixed() + '</Data></Cell>' ;
						}
						if(Fe ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +='<Cell><Data ss:Type="String">$' + c15.toFixed() + '</Data></Cell>' ;
						}
						if(Mr ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=		'<Cell><Data ss:Type="String">$' + c16.toFixed() + '</Data></Cell>' ;
						}
						if(Ap ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=		'<Cell><Data ss:Type="String">$' + c17.toFixed() + '</Data></Cell>' ;
						}
						if(Ma ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=		'<Cell><Data ss:Type="String">$' + c18.toFixed() + '</Data></Cell>' ;
						}
						if(Ju ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
						excelObj  +=	'<Cell><Data ss:Type="String">$' + c19.toFixed() + '</Data></Cell>' ;
						}
						if(Jl ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=		'<Cell><Data ss:Type="String">$' + c20.toFixed() + '</Data></Cell>' ;
						}
						if(Au ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
						excelObj  +=	'<Cell><Data ss:Type="String">$' + c21.toFixed() + '</Data></Cell>' ;
						}
						if(Se ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
						excelObj  +=	'<Cell><Data ss:Type="String">$' + c22.toFixed() + '</Data></Cell>' ;
						}
						if(Oc ==(Number(m_Date)-2))
						{
						excelObj  +='<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=		'<Cell><Data ss:Type="String">$' + c23.toFixed() + '</Data></Cell>' ;
						}
						if(No ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
						excelObj  +=	'<Cell><Data ss:Type="String">$' + c24.toFixed() + '</Data></Cell>' ;
						}
						if(De ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalCrrMonth.toFixed() + '</Data></Cell>' ;
						}else{
						excelObj  +=	'<Cell><Data ss:Type="String">$' + c38.toFixed() + '</Data></Cell>' ;
						}
						/*'<Cell><Data ss:Type="String">$' + c14.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c15.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c16.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c17.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c18.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c19.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c20.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c21.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c22.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c23.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c24.toFixed() + '</Data></Cell>' +
						'<Cell><Data ss:Type="String">$' + c38.toFixed() + '</Data></Cell>' +
						*/
						if(Ja ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c25.toFixed()+ '</Data></Cell>' ;
						}
						if(Fe ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c26.toFixed()+'</Data></Cell>' ;
						}
						if(Mr ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c27.toFixed()+'</Data></Cell>' ;
						}
						if(Ap ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c28.toFixed()+'</Data></Cell>' ;
						}
						if(Ma ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c29.toFixed()+'</Data></Cell>' ;
						}
						if(Ju ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c30.toFixed()+'</Data></Cell>' ;
						}
						if(Jl ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c31.toFixed()+ '</Data></Cell>' ;
						}
						if( Au==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c32.toFixed()+ '</Data></Cell>';
						}
						if(Se ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c33.toFixed()+ '</Data></Cell>';
						}
						if(Oc ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c34.toFixed()+ '</Data></Cell>';
						}
						if(No ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c36.toFixed()+ '</Data></Cell>';
						}
						if(De ==(Number(m_Date)-2))
						{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + subTotalOppMonth.toFixed() + '</Data></Cell>' ;
						}else{
					excelObj  +=	'<Cell><Data ss:Type="String">$' + c37.toFixed()+ '</Data></Cell>' ;
						}
					excelObj  +=	'</Row>';
		}
		else {
			htmlObj1 +="<tr>";
			htmlObj1 +="<td colspan='15'>No records to show.</td>";
			htmlObj1 +="</tr>";
			
			excelObj += '<Row>'+'<Cell><Data ss:Type="String">No records to show.</Data></Cell></Row>';
		}
		htmlObj1 +='</table>';
		log.debug('totalValue',totalValue)
		//log.debug({title: "htmlObj1", details: htmlObj1});
		//log.debug({title: "excelObj", details: excelObj});
		var finalString  = htmlObj1 + ":||:" + excelObj;
		//log.debug({title: "finalString", details: finalString});
		return finalString;
		//context.response.write({output: finalString });

	}
	
	
		function setMonthYearData(monthRange, yearRange) {
			var month;
			var currentDate = new Date();
			var monthNum = currentDate.getMonth();
			var yearText = currentDate.getFullYear();
			var previousYear = yearText;
			monthRange.addSelectOption({
							value:'',
							text:''
						})
			for(var i = 0; i < 12; i++) {
				switch(i) {
					case 0:
						month = 'Jan';
						break;
					case 1:
						month = 'Feb';
						break;
					case 2:
						month = 'Mar';
						break;
					case 3:
						month = 'Apr';
						break;
					case 4:
						month = 'May';
						break;
					case 5:
						month = 'June';
						break;
					case 6:
						month = 'July';
						break;
					case 7:
						month = 'Aug';
						break;
					case 8:
						month = 'Sep';
						break;
					case 9:
						month = 'Oct';
						break;
					case 10:
						month = 'Nov';
						break;
					case 11:
						month = 'Dec'

				}
			/*	if(monthNum == i) {
					monthRange.addSelectOption({
						value: i,
						text: month,
						isSelected: true
					});
				}
				else */{
					monthRange.addSelectOption({
						value: i,
						text: month
						//isSelected : true
					});
				}

			} //end for(var i=0;i<12;i++) 
			for(var k = 0; k < 3; k++) {
				yearRange.addSelectOption({
					value: Number(yearText - k),
					text: (yearText - k),
					//isSelected : true
				});
			}

		}
		function _logValidation(value) {
			if(value != 'null' && value != null && value != null && value != '' && value != undefined && value != undefined && value != 'undefined' && value != 'undefined' && value != 'NaN' && value != NaN && value!='Infinity') {
				return true;
			}
			else {
				return false;
			}
		}
		function clientPoNumner(gstin)
		{
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
			gstin.addSelectOption({
				value:result.getValue({ name: "custentity_mp_client_po_num",
         summary: "GROUP",
         label: "Client PO Number"}),
				text:result.getValue({ name: "custentity_mp_client_po_num",
         summary: "GROUP",
         label: "Client PO Number"})
			})
		   // .run().each has a limit of 4,000 results
		   return true;
		});
		}
function serchDataByClient(mum_value,gstCustomerId,yearId)
{
	var customrecord_sss_proj_snapshotSearchObj = search.create({
				type: "customrecord_sss_proj_snapshot",
				filters: [
					["custrecord_sss_ss_run.custrecord_sss_run_deployment", "contains", "snap4"],
					"AND",
					["custrecord_sss_ss_run.custrecord_sss_run_is_current", "is", "T"],
					"AND",
					["custrecord_sss_ss_proj.customer", "anyof", gstCustomerId],
					"AND",
					["custrecord_sss_ss_proj.custentity_pixacore_project_year", "is", yearId],
					"AND",
					["custrecord_sss_ss_proj.custentity_mp_client_po_num", "is",mum_value]

				],
				columns: [
					search.createColumn({
						name: "formulatext",
						formula: "case when to_char({custrecord_sss_ss_proj.customer}) = substr(to_char({custrecord_sss_ss_proj.altname}),1,REGEXP_INSTR(to_char({custrecord_sss_ss_proj.altname}),' : ')-1) then to_char({custrecord_sss_ss_proj.customer}) else concat(substr(to_char({custrecord_sss_ss_proj.altname}),1,REGEXP_INSTR(to_char({custrecord_sss_ss_proj.altname}),' : ')+2),to_char({custrecord_sss_ss_proj.customer})) end",
						sort: search.Sort.ASC,
						label: "Client : Brand"
					}),
					search.createColumn({
						name: "formulatext",
						formula: "case when {custrecord_sss_ss_proj.custentity_mp_is_sow} = 'T' then {custrecord_sss_ss_proj.entityid} else concat(nvl(substr({custrecord_sss_ss_proj.custentity_mp_sow_project},1,7),'[sow]'),concat(', ',concat(nvl({custrecord_sss_ss_proj.custentity_mp_client_po_num},'[po#]'),concat(', ',(case when {custrecord_sss_ss_proj.custentity_mp_parent_project.id} is null then {custrecord_sss_ss_proj.entityid} else concat(substr({custrecord_sss_ss_proj.custentity_mp_parent_project},1,7),concat(', ',{custrecord_sss_ss_proj.entityid})) end))))) end",
						sort: search.Sort.ASC,
						label: "SOW, PO Number, Project Num"
					}),
					search.createColumn({
						name: "companyname",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Project Name"
					}),
					search.createColumn({
						name: "entityid",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Project Number"
					}),
					search.createColumn({
						name: "startdate",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Start Date"
					}),
					search.createColumn({
						name: "projectedenddate",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Projected End Date"
					}),
					search.createColumn({
						name: "custentity_pixacore_project_year",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Project Year"
					}),
					search.createColumn({
						name: "entitystatus",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Status"
					}),
					search.createColumn({
						name: "custentity_mp_proj_sow_fee",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "SOW Fee Budget"
					}),
					search.createColumn({
						name: "custentity_mp_proj_sow_oop",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "SOW OOP Budget"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_est3},0) + nvl({custrecord_sss_ss_proj.custentity_mp_prj_est_fee},0)",
						label: "Estimated Fees"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_est1},0) +  nvl({custrecord_sss_ss_proj.custentity_mp_prj_est_oops},0)                           ",
						label: "Estimated OOP"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_est3},0) + nvl({custrecord_sss_ss_proj.custentity_mp_prj_est_fee},0) + 	nvl({custrecord_sss_ss_est1},0) + nvl({custrecord_sss_ss_proj.custentity_mp_prj_est_oops},0)",
						label: "Estimate Total"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_time2},0) + nvl({custrecord_sss_ss_proj.custentity_mp_prj_actualfee},0)",
						label: "Actual Fee"
					}),
					search.createColumn({
						name: "custrecord_sss_ss_cost1",
						label: "Actual OOP"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_time2},0) + nvl({custrecord_sss_ss_proj.custentity_mp_prj_actualfee},0) + nvl({custrecord_sss_ss_cost1},0)",
						label: "Actual Total"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_tran2},0) + nvl({custrecord_sss_ss_tran3},0) + nvl({custrecord_sss_ss_tran4},0)  ",
						label: "Billed to Date"
					}),
					search.createColumn({
						name: "custrecord_sss_ss_time5",
						label: "New Billing Fee"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_tran6},0) + nvl({custrecord_sss_ss_tran7},0) ",
						label: "New Billing OOP Med"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_tran5},0) +nvl({custrecord_sss_ss_tran6},0) + nvl({custrecord_sss_ss_tran7},0) ",
						label: "New Billing Amt"
					}),
					search.createColumn({
						name: "formulacurrency",
						formula: "nvl({custrecord_sss_ss_tran2},0) + nvl({custrecord_sss_ss_tran3},0) + nvl({custrecord_sss_ss_tran4},0)  + nvl({custrecord_sss_ss_tran5},0) +nvl({custrecord_sss_ss_tran6},0) + nvl({custrecord_sss_ss_tran7},0) ",
						label: "Total Billing with New"
					}),
					search.createColumn({
						name: "custrecord_sss_ss_cost5",
						label: "Third Party Cost"
					}),
					search.createColumn({
						name: "custrecord_sss_ss_last_date",
						label: "Last Activity Date"
					}),
					search.createColumn({
						name: "custrecord_sss_compl_date",
						join: "CUSTRECORD_SSS_SS_RUN",
						label: "Snapshot Dt Time"
					}),
					search.createColumn({
						name: "subsidiarynohierarchy",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Subsidiary (no hierarchy)"
					}),
					search.createColumn({
						name: "custentity_mp_client_po_num",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Client PO Number"
					})
				]
			});
			var htmlStr1 = "";
			var total_Sow_Opp = 0;
			var customrecord_sss_proj_snapshot = [];
			var project_Number_Array = [];
			var searchResultCount = customrecord_sss_proj_snapshotSearchObj.runPaged()
				.count;
			log.debug("jobSearchObj result count", searchResultCount);
			customrecord_sss_proj_snapshotSearchObj.run()
				.each(function(result) {
					var Client_PO_Number = result.getValue(result.columns[1]); //Client PO Number
					log.debug('Client_PO_Number',Client_PO_Number)
					var Project_Number = result.getValue(result.columns[3]); //Project Number
					log.debug('Project_Number',Project_Number)
					project_Number_Array.push(Project_Number);
					var Client_Project = result.getValue(result.columns[2]); //Client : Project
					var Status = result.getText({
						name: "entitystatus",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Status"
					});
					var Start_Date = result.getValue(result.columns[4]);
					var Projected_End_Date = result.getValue(result.columns[5]);

					var sowFee = result.getValue(result.columns[8]);
					if(sowFee) {
						sowFee = sowFee;
					}
					else {
						sowFee = 0
					}
					var sowOpp = result.getValue(result.columns[9]);
					if(sowOpp) {
						sowOpp = sowOpp;
					}
					else {
						sowOpp = 0;
					}
					total_Sow_Opp = Number(sowOpp) + Number(sowFee)


					var ACTUAL_FEE = result.getValue(result.columns[13]);
					var ACTUAL_OOP = result.getValue(result.columns[14]);
					var actual_TOTAL = result.getValue(result.columns[15]);
					var ESTIMATED_FEES = result.getValue(result.columns[10]);
					var ESTIMATED_OOP = result.getValue(result.columns[11]);
					var ESTIMATE_TOTAL = result.getValue(result.columns[12]);
					
					var Over_Under = Math.abs(Number(ESTIMATE_TOTAL)-Number(total_Sow_Opp))
					var over_under = Math.abs(Number(ESTIMATE_TOTAL)-Number(actual_TOTAL))
					if(actual_TOTAL && ESTIMATE_TOTAL)
					{
						var newValue =Number( actual_TOTAL)/Number(ESTIMATE_TOTAL)
						var utilized = newValue*100;
					}
					
					
					var Billed_to_Date = result.getValue(result.columns[16]);
					var cli_po_num = result.getValue({
						name: "custentity_mp_client_po_num",
						join: "CUSTRECORD_SSS_SS_PROJ",
						label: "Client PO Number"
					});
					log.debug('cli_po_num', cli_po_num)
					jsonAyyaData.push({
						"Client_PO_Number":Client_PO_Number,
						"Project_Number":Project_Number,
						"Client_Project":Client_Project,
						"Status":Status,
						"Start_Date":Start_Date,
						"Projected_End_Date":Projected_End_Date,
						"sowFee":sowFee,
						"total_Sow_Opp":total_Sow_Opp,
						"ACTUAL_FEE":ACTUAL_FEE,
						"ACTUAL_OOP":ACTUAL_OOP,
						"actual_TOTAL":actual_TOTAL,
						"ESTIMATED_FEES":ESTIMATED_FEES,
						"ESTIMATED_OOP":ESTIMATED_OOP,
						"ESTIMATE_TOTAL":ESTIMATE_TOTAL,
						"Over_Under":Over_Under,
						"over_under":over_under,
						"utilized":utilized,
						"Billed_to_Date":Billed_to_Date,
						"cli_po_num":cli_po_num,
						"yearValue":yearId

					})
					return true;
				});
				log.debug('jsonAyyaData',jsonAyyaData)
}
function secordSearch(gstCustomerId,cli_po_num,Project_Number,year_Value)
{
	var monthlyArray = [];
					var jobSearchObj = search.create({
						type: "job",
						filters: [
							/*["transaction.type", "anyof", "SalesOrd"],
							"AND",
							["transaction.trandate", "within", "thisyear"],
							"AND",
							["transaction.custbody_mp_sow_bill_status", "anyof", "1", "2", "3"],
							"AND",
							["transaction.status", "noneof", "SalesOrd:C"]*/
							["transaction.type", "anyof", "SalesOrd"],
							"AND",
							["transaction.trandate", "within", "thisyear"],
							"AND",
							["transaction.custbody_mp_sow_bill_status", "anyof", "1", "2", "3"],
							"AND",
							["transaction.status", "noneof", "SalesOrd:C"],
							"AND",
							["customer", "anyof", gstCustomerId],
							"AND",
							["custentity_pixacore_project_year", "is", year_Value],
							"AND",
							["custentity_mp_client_po_num", "is", cli_po_num],
							"AND", 
							["entityid","is",Project_Number]
						],
						columns: [
							search.createColumn({
								name: "customer",
								summary: "GROUP",
								label: "Client"
							}),
							search.createColumn({
								name: "custentity_mp_client_po_num",
								summary: "GROUP",
								label: "Client PO Number"
							}),
							search.createColumn({
								name: "entityid",
								summary: "GROUP",
								sort: search.Sort.ASC,
								label: "Project Number"
							}),
							search.createColumn({
								name: "companyname",
								summary: "GROUP",
								label: "Project Name"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2101' then {transaction.amount} else null end",
								label: "January"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2102' then {transaction.amount} else null end",
								label: "February"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2103' then {transaction.amount} else null end",
								label: "March"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2104' then {transaction.amount} else null end",
								label: "April"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2105'  then {transaction.amount} else null end",
								label: "May"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2106' then {transaction.amount} else null end",
								label: "June"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2107' then {transaction.amount} else null end",
								label: "July"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2108' then {transaction.amount} else null end",
								label: "August"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2109' then {transaction.amount} else null end",
								label: "September"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2110' then {transaction.amount} else null end",
								label: "October"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2111' then {transaction.amount} else null end",
								label: "November"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} <> 'Sales Order' then null when to_char({transaction.trandate},'YYMM') = '2112' then {transaction.amount} else null end",
								label: "December"
							}),
							search.createColumn({
								name: "formulacurrency",
								summary: "SUM",
								formula: "case when {transaction.type} = 'Sales Order' then {transaction.amount} else null end",
								label: "Total"
							}),
							search.createColumn({
								name: "item",
								join: "transaction",
								summary: "GROUP",
								label: "Item"
							})
						]
					});
					var newAr = [];
					var anoth = [];
					var obj = {
						key1: []
					};
					var searchResultCount = jobSearchObj.runPaged()
						.count;
					log.debug("jobSearchObj result count", searchResultCount);
					jobSearchObj.run()
						.each(function(result) {
							
						//	var projID=result.getValue(result.columns[2]),
							var January = result.getValue(result.columns[4]);
							var February = result.getValue(result.columns[5]);
							var March = result.getValue(result.columns[6]);
							var April = result.getValue(result.columns[7]);
							var May = result.getValue(result.columns[8]);
							var June = result.getValue(result.columns[9]);
							var July = result.getValue(result.columns[10]);
							var August = result.getValue(result.columns[11]);
							var September = result.getValue(result.columns[12]);
							var October = result.getValue(result.columns[13]);
							var November = result.getValue(result.columns[14]);
							var December = result.getValue(result.columns[15]);
							monthlyArray.push({
								"projectID": result.getValue(result.columns[2]),
								"itemType": result.getText({
									name: "item",
									join: "transaction",
									summary: "GROUP",
									label: "Item"
								}),
								"January": January,
								"February": February,
								"March": March,
								"April": April,
								"May": May,
								"June": June,
								"July": July,
								"August": August,
								"September": September,
								"October": October,
								"November": November,
								"December": December,

							})
							// .run().each has a limit of 4,000 results
							return true;
						});
					log.debug('monthlyArray', monthlyArray.length);
	return monthlyArray
}
function thirdSerchForOpp(gstCustomerId,cli_po_num,Project_Number,year_Value)
{
	var oppMonthArray = [];
					var customrecord_mp_proj_forecastsSearchObj = search.create({
						type: "customrecord_mp_proj_forecasts",
						filters: [
							["custrecord_mp_projfore_type", "anyof", "1", "2"],
							"AND",
							["custrecord_mp_projfore_proj.custentity_pixacore_project_year", "is", year_Value],
							"AND",
							["custrecord_mp_projfore_proj.customer", "anyof",gstCustomerId],
							"AND",
							["custrecord_mp_projfore_proj.custentity_mp_client_po_num", "is", cli_po_num],
							"AND", 
							["custrecord_mp_projfore_proj.entityid","is",Project_Number]
						],
						columns: [
							search.createColumn({
								name: "entityid",
								join: "CUSTRECORD_MP_PROJFORE_PROJ",
								label: "ID"
							}),
							search.createColumn({
								name: "custentity_cca_select_client",
								join: "CUSTRECORD_MP_PROJFORE_PROJ",
								label: "Select Client for Project"
							}),
							search.createColumn({
								name: "companyname",
								join: "CUSTRECORD_MP_PROJFORE_PROJ",
								label: "Project Name"
							}),
							search.createColumn({
								name: "custentity_pixacore_project_year",
								join: "CUSTRECORD_MP_PROJFORE_PROJ",
								label: "Project Year"
							}),
							search.createColumn({
								name: "entitystatus",
								join: "CUSTRECORD_MP_PROJFORE_PROJ",
								label: "Status"
							}),
							search.createColumn({
								name: "formulatext",
								formula: "concat({custrecord_mp_projfore_year},concat('  ',{custrecord_mp_projfore_type}))",
								label: "Year Type"
							}),
							search.createColumn({
								name: "custrecord_mp_projfore_type",
								sort: search.Sort.ASC,
								label: "Line Type"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon1",
								label: "January"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon2",
								label: "February"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon3",
								label: "March"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon4",
								label: "April"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon5",
								label: "May"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon6",
								label: "June"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon7",
								label: "July"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon8",
								label: "August"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon9",
								label: "September"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon10",
								label: "October"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon11",
								label: "November"
							}),
							search.createColumn({
								name: "custrecord_mp_pf_mon12",
								label: "December"
							}),
							search.createColumn({
								name: "formulacurrency",
								formula: "{custrecord_mp_pf_mon1}+{custrecord_mp_pf_mon2}+{custrecord_mp_pf_mon3}+{custrecord_mp_pf_mon4}+{custrecord_mp_pf_mon5}+{custrecord_mp_pf_mon6}+{custrecord_mp_pf_mon7}+{custrecord_mp_pf_mon8}+{custrecord_mp_pf_mon9}+{custrecord_mp_pf_mon10}+{custrecord_mp_pf_mon11}+{custrecord_mp_pf_mon12}",
								label: "Total"
							}),
							search.createColumn({
								name: "formulatext",
								formula: "{custrecord157}",
								label: "Project"
							}),
							search.createColumn({
								name: "formulatext",
								formula: "{custrecord_mp_projfore_proj.altname}",
								label: "Client"
							}),
							search.createColumn({
								name: "formulatext",
								formula: "{custrecord_mp_projfore_proj.entityid} ",
								label: "Project Number"
							}),
							search.createColumn({
								name: "formulatext",
								formula: "{custrecord_mp_projfore_proj.jobname}",
								label: "Project Name"
							})
						]
					});

					var searchResultCount = customrecord_mp_proj_forecastsSearchObj.runPaged()
						.count;
					log.debug("jobSearchObj result count", searchResultCount);
					customrecord_mp_proj_forecastsSearchObj.run()
						.each(function(result) {
							//for(var  i=0;i<salesResults.length;i++)
							

								oppMonthArray.push({
									"projectID": result.getValue(result.columns[0]),
									"itemType": result.getValue(result.columns[6]),
									"January": result.getValue(result.columns[7]),
									"February": result.getValue(result.columns[8]),
									"March": result.getValue(result.columns[9]),
									"April": result.getValue(result.columns[10]),
									"May": result.getValue(result.columns[11]),
									"June": result.getValue(result.columns[12]),
									"July": result.getValue(result.columns[13]),
									"August": result.getValue(result.columns[14]),
									"September": result.getValue(result.columns[15]),
									"October": result.getValue(result.columns[16]),
									"November": result.getValue(result.columns[17]),
									"December": result.getValue(result.columns[18]),
								})
							return true;
							
						});
					log.debug('oppMonthArray in search', oppMonthArray)
					return  oppMonthArray
}
	return {
		onRequest : onRequest
	}
});