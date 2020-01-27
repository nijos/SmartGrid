import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { string } from "prop-types";
type DataSet = ComponentFramework.PropertyTypes.DataSet;
// Define const here
const RowRecordId:string = "rowRecId";
//list of common read only fields
let readOnly: string[] ;
readOnly= ['createdby', 'createdonbehalfby', 'createdbyexternalparty','createdon','processid','statecode','statuscode'];

// Style name of Load More Button
const LoadMoreButton_Hidden_Style = "LoadMoreButton_Hidden_Style";
export class SmartGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	// Cached context object for the latest updateView
	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;

	// Table element created as part of this control's table
	private dataTable: HTMLTableElement;

	// Button element created as part of this control
	private loadPageButton: HTMLButtonElement;
	 

	
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Need to track container resize so that control could get the available width. The available height won't be provided even this is true
		context.mode.trackContainerResize(true);

		// Create main table container div. 
		this.mainContainer = document.createElement("div");
		this.mainContainer.classList.add("SimpleTable_MainContainer_Style");

		// Create data table container div. 
		this.dataTable = document.createElement("table");
		this.dataTable.classList.add("SimpleTable_Table_Style");

		// Create data table container div. 
		this.loadPageButton = document.createElement("button");
		this.loadPageButton.setAttribute("type", "button");
		this.loadPageButton.innerText = context.resources.getString("PCF_TSTableGrid_LoadMore_ButtonLabel");
		this.loadPageButton.classList.add(LoadMoreButton_Hidden_Style);
		this.loadPageButton.classList.add("LoadMoreButton_Style");
		this.loadPageButton.addEventListener("click", this.onLoadMoreButtonClick.bind(this));

		// Adding the main table and loadNextPage button created to the container DIV.
		this.mainContainer.appendChild(this.dataTable);
		this.mainContainer.appendChild(this.loadPageButton);
		container.appendChild(this.mainContainer);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.contextObj = context;
			this.toggleLoadMoreButtonWhenNeeded(context.parameters.smartGridDataSet);

			if(!context.parameters.smartGridDataSet.loading){
				
				// Get sorted columns on View
				let columnsOnView = this.getSortedColumnsOnView(context);

				if (!columnsOnView || columnsOnView.length === 0) {
                    return;
				}

				let columnWidthDistribution = this.getColumnWidthDistribution(context, columnsOnView);


				while(this.dataTable.firstChild)
				{
					this.dataTable.removeChild(this.dataTable.firstChild);
				}

				this.dataTable.appendChild(this.createTableHeader(columnsOnView, columnWidthDistribution));		
				this.dataTable.appendChild(this.createTableBody(columnsOnView, columnWidthDistribution, context.parameters.smartGridDataSet));
				this.dataTable.appendChild(this.createNewSection(columnsOnView, columnWidthDistribution));
				this.dataTable.parentElement!.style.height = window.innerHeight - this.dataTable.offsetTop - 70 + "px";
			
			}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

		/**
		 * Get sorted columns on view
		 * @param context 
		 * @return sorted columns object on View
		 */
		private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[]
		{
			if (!context.parameters.smartGridDataSet.columns) {
				return [];
			}
			
			let columns =context.parameters.smartGridDataSet.columns
				.filter(function (columnItem:DataSetInterfaces.Column) { 
					// some column are supplementary and their order is not > 0
					return columnItem.order >= 0 }
				);
			
			// Sort those columns so that they will be rendered in order
			columns.sort(function (a:DataSetInterfaces.Column, b: DataSetInterfaces.Column) {
				return a.order - b.order;
			});


			
			return columns;
		}
			/**
		 * Get column width distribution
		 * @param context context object of this cycle
		 * @param columnsOnView columns array on the configured view
		 * @returns column width distribution
		 */
		private getColumnWidthDistribution(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[]): string[]{

			let widthDistribution: string[] = [];
			
			// Considering need to remove border & padding length
			let totalWidth:number = context.mode.allocatedWidth - 250;
			let widthSum = 0;
			
			columnsOnView.forEach(function (columnItem) {
				widthSum += columnItem.visualSizeFactor;
			});

			let remainWidth:number = totalWidth;
			
			columnsOnView.forEach(function (item, index) {
				let widthPerCell = "";
				if (index !== columnsOnView.length - 1) {
					let cellWidth = Math.round((item.visualSizeFactor / widthSum) * totalWidth);
					remainWidth = remainWidth - cellWidth;
					widthPerCell = cellWidth + "px";
				}
				else {
					widthPerCell = remainWidth + "px";
				}
				widthDistribution.push(widthPerCell);
			});

			return widthDistribution;

		}

		private createTableHeader(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[]):HTMLTableSectionElement{

			let tableHeader:HTMLTableSectionElement = document.createElement("thead");
			let tableHeaderRow: HTMLTableRowElement = document.createElement("tr");
			tableHeaderRow.classList.add("SimpleTable_TableRow_Style");

			columnsOnView.forEach(function(columnItem, index){
				
				let tableHeaderCell = document.createElement("th");
				tableHeaderCell.classList.add("SimpleTable_TableHeader_Style");
				let innerDiv = document.createElement("div");
				innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
				innerDiv.style.maxWidth = widthDistribution[index];
				innerDiv.innerText = columnItem.displayName;
				tableHeaderCell.appendChild(innerDiv);
				tableHeaderRow.appendChild(tableHeaderCell);
			});

			tableHeader.appendChild(tableHeaderRow);
			return tableHeader;
		}


//section for create new 
private createNewSection(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[]):HTMLDivElement{
	var thisRef = this;
    let filterContainer: HTMLDivElement = document.createElement("div");
    filterContainer.classList.add("col-md-12");
	
	columnsOnView.forEach(function(columnItem, index){
		if(columnItem.dataType=="OptionSet")
		{
			
			let _select:HTMLSelectElement;
			_select=document.createElement("select");
			_select.setAttribute("id", "in_" + columnItem.name);
			_select.setAttribute("name",columnItem.dataType);
			// get optionset metadata
					
			var req1 = new XMLHttpRequest();
						
			req1.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.0/EntityDefinitions(LogicalName='"+thisRef.contextObj.parameters.smartGridDataSet.getTargetEntityType()+"')/Attributes(LogicalName='"+columnItem.name+"')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)", false);
			req1.setRequestHeader("OData-MaxVersion", "4.0");			
			req1.setRequestHeader("OData-Version", "4.0");
			req1.setRequestHeader("Accept", "application/json");
			req1.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req1.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
			req1.onreadystatechange = function() {
			if (this.readyState === 4) {
				req1.onreadystatechange = null;
				if (this.status === 200) {
					var resultdata = JSON.parse(this.response);
					if((resultdata.OptionSet!=null)&&(resultdata.OptionSet!="")&&(resultdata.OptionSet!="undefined"))
				{
					  
					// @ts-ignore 
					resultdata.OptionSet.Options.forEach(function(option) {
					var optionitem:HTMLOptionElement = document.createElement("option");
						// @ts-ignore 
					optionitem.value=option!.Value;
						// @ts-ignore 
					optionitem.text = option!.Label.UserLocalizedLabel.Label;
						// @ts-ignore 
					_select.add(optionitem);



				});
				
				}
				} 
				else {
					alert("smartgrid:03: and error occured while getting entity metadata for options"+this.statusText);
				}
			}
		};
		req1.send();

			//retrived optionset medata
			
			_select.classList.add("selectstyle");
			filterContainer.appendChild(_select);
		}
		else if(columnItem.dataType=="TwoOptions")
		{
		
			let innerCheckbox = document.createElement("input");
			innerCheckbox.setAttribute("id", "in_" + columnItem.name);
			innerCheckbox.setAttribute("type", "checkbox");
			innerCheckbox.setAttribute("name",columnItem.dataType);
			innerCheckbox.classList.add("box");
			innerCheckbox.checked = false;
			let innerLabel: HTMLLabelElement = document.createElement("label");
                        innerLabel.classList.add("Textbox_Style");
                        innerLabel.setAttribute("for", innerCheckbox.id);
                        innerLabel.innerHTML= columnItem.displayName;
                        innerLabel.setAttribute("name", "label_"+columnItem.name);
                        filterContainer.appendChild(innerLabel);
			            filterContainer.appendChild(innerCheckbox);	
		}
		
		else
		{

		let _Topic:HTMLInputElement;
		_Topic = document.createElement("input");	
	
		_Topic.setAttribute("id", "in_" + columnItem.name);
		_Topic.setAttribute("name",columnItem.dataType);
		_Topic.placeholder = columnItem.displayName;
		_Topic.classList.add("Textbox_Style");
		_Topic.style.width = widthDistribution[index];
		_Topic.style.display = "inline-block";
		// lock if the field is non editable
		if(readOnly.includes(columnItem.name))
		{
			_Topic.readOnly=true;
			_Topic.classList.add("readonlyInput");
		}
	
		if(columnItem.dataType.startsWith("Lookup"))
		{
			_Topic.readOnly=true;
			_Topic.placeholder = "Select "+columnItem.displayName;
			if((columnItem.name!=thisRef.contextObj.parameters.primaryLookupName.raw)&&(columnItem.dataType!="Lookup.Customer"))
			_Topic.addEventListener("click",( e:Event)=>thisRef.openLookupDialogue(_Topic.id));
			//Create a hidden field to capture lookup values
									let recordGuid:HTMLInputElement;
									recordGuid = document.createElement("input");	
									recordGuid.setAttribute("type", "hidden");
									recordGuid.setAttribute("id", "hd_in_"+columnItem.name);
									recordGuid.setAttribute("name",columnItem.dataType);
									filterContainer.appendChild(recordGuid);

		}
		
		else if(columnItem.dataType.includes("DateAndTime"))
		{
			_Topic.setAttribute("type", "date");
		}
		else if(columnItem.dataType.includes("Decimal"))
		{
			_Topic.setAttribute("type", "number");
			_Topic.setAttribute("step", "1.00");
			_Topic.setAttribute("min", "0.00");
			
		}
		else if(columnItem.dataType.includes("Whole"))
		{
			_Topic.setAttribute("type", "number");
		}
	
	
		else
		{
		_Topic.setAttribute("type", "text");
		}
		filterContainer.appendChild(_Topic);
	}



		
	});

	let addNew:HTMLButtonElement;
    addNew = document.createElement("button");
    addNew.setAttribute("type", "button");
    addNew.innerText = "+";
    addNew.classList.add("butn");
    addNew.classList.add("butn:hover");
  	addNew.addEventListener("click", this.addNewButtonClick.bind(this));
	filterContainer.appendChild(addNew);
	
	return filterContainer;
}
//
		private createTableBody(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[], gridParam: DataSet):HTMLTableSectionElement{

			let tableBody:HTMLTableSectionElement = document.createElement("tbody");

			if(gridParam.sortedRecordIds.length > 0)
			{
				for(let currentRecordId of gridParam.sortedRecordIds){
					
					let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
					tableRecordRow.classList.add("SimpleTable_TableRow_Style");
					tableRecordRow.addEventListener("click", this.onRowClick.bind(this));

					// Set the recordId on the row dom
					tableRecordRow.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());
					

					columnsOnView.forEach(function(columnItem, index){
						let tableRecordCell = document.createElement("td");
						tableRecordCell.classList.add("SimpleTable_TableCell_Style");
						let innerDiv = document.createElement("div");
						innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
						innerDiv.style.maxWidth = widthDistribution[index];
						innerDiv.innerText = gridParam.records[currentRecordId].getFormattedValue(columnItem.name);
						tableRecordCell.appendChild(innerDiv);
						tableRecordRow.appendChild(tableRecordCell);
					});

					tableBody.appendChild(tableRecordRow);
				}
			}
			else
			{
				let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
				let tableRecordCell: HTMLTableCellElement = document.createElement("td");
				tableRecordCell.classList.add("No_Record_Style");
				tableRecordCell.colSpan = columnsOnView.length;
				tableRecordCell.innerText = this.contextObj.resources.getString("PCF_TSTableGrid_No_Record_Found");
				tableRecordRow.appendChild(tableRecordCell)
				tableBody.appendChild(tableRecordRow);
			}

			return tableBody;
		}

			/**
		 * Row Click Event handler for the associated row when being clicked
		 * @param event
		 */
		private onRowClick(event: Event): void {
			let rowRecordId = (event.currentTarget as HTMLTableRowElement).getAttribute(RowRecordId);

			if(rowRecordId)
			{
				let gridEntity: string=this.contextObj.parameters.smartGridDataSet.getTargetEntityType().toString();
				let entityReference = this.contextObj.parameters.smartGridDataSet.records[rowRecordId].getNamedReference();
				let entityFormOptions = {
					entityName: gridEntity,
					entityId: entityReference.id.toString()
				}
				this.contextObj.navigation.openForm(entityFormOptions);
			}
		}

		/**
		 * Toggle 'LoadMore' button when needed
		 */
		private toggleLoadMoreButtonWhenNeeded(gridParam: DataSet): void{
			
			if(gridParam.paging.hasNextPage && this.loadPageButton.classList.contains(LoadMoreButton_Hidden_Style))
			{
				this.loadPageButton.classList.remove(LoadMoreButton_Hidden_Style);
			}
			else if(!gridParam.paging.hasNextPage && !this.loadPageButton.classList.contains(LoadMoreButton_Hidden_Style))
			{
				this.loadPageButton.classList.add(LoadMoreButton_Hidden_Style);
			}

		}
	/**
		 * 'LoadMore' Button Event handler when load more button clicks
		 * @param event
		 */
		private onLoadMoreButtonClick(event: Event): void {
			this.contextObj.parameters.smartGridDataSet.paging.loadNextPage();
			this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.smartGridDataSet);
			
		}

		private addNewButtonClick (event:Event): void {
			
			let columnsOnView :DataSetInterfaces.Column[]= this.getSortedColumnsOnView(this.contextObj);
			let data: any = {};
			var thisRef= this;
			columnsOnView.forEach(function (columnItem, index) {

				 if(columnItem.dataType=="OptionSet")
				{
					var fieldOptionSet = document.getElementById("in_" + columnItem.name)as HTMLSelectElement;
					let selectedOption:number=fieldOptionSet.selectedIndex;
					data[columnItem.name]=fieldOptionSet.options[selectedOption].value;
				}
			else if(columnItem.dataType=="Decimal")
				{
					var fieldDecimal = document.getElementById("in_" + columnItem.name)as HTMLInputElement;
					if (fieldDecimal.value != null && fieldDecimal.value != "" && fieldDecimal.value != "undefined") {

						data[columnItem.name] = Number( fieldDecimal.value);
					}
				
				}
				else if(columnItem.dataType=="TwoOptions")
				{
				
						let checkBox: HTMLInputElement = <HTMLInputElement>document.getElementById("in_" + columnItem.name);
						data[columnItem.name] =( (checkBox.checked == true) ? true : false); //in the onclick it's still the old value which is being switched
						
					
				}
			  else if(columnItem.dataType.startsWith("Lookup") && columnItem.dataType!="Lookup.Customer" )
			  {
				var field = document.getElementById("in_" + columnItem.name)as HTMLInputElement;
				var req = new XMLHttpRequest();
				req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.0/EntityDefinitions(LogicalName='"+thisRef.contextObj.parameters.smartGridDataSet.getTargetEntityType()+"')/Attributes(LogicalName='"+columnItem.name+"')", false);
				req.setRequestHeader("OData-MaxVersion", "4.0");
				req.setRequestHeader("OData-Version", "4.0");
				req.setRequestHeader("Accept", "application/json");
				req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
				req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
				req.onreadystatechange = function() {
					if (this.readyState === 4) {
						req.onreadystatechange = null;
						if (this.status === 200) {
							var result = JSON.parse(this.response);
							if((result.Targets[0]!=null)&&(result.Targets[0]!="")&&(result.Targets[0]!="undefined") )
						{
  
							let  targetName: string= result.Targets[0];
							let schemaName:string=result.LogicalName!;
						//	data[schemaName+"@odata.bind"] = "/"+targetName+"s(" + thisRef.contextObj.page.entityId + ")";
						  
					
							var req1 = new XMLHttpRequest();
						
							req1.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.0/EntityDefinitions(LogicalName='"+targetName+"')", false);
							req1.setRequestHeader("OData-MaxVersion", "4.0");
							req1.setRequestHeader("OData-Version", "4.0");
							req1.setRequestHeader("Accept", "application/json");
							req1.setRequestHeader("Content-Type", "application/json; charset=utf-8");
							req1.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
							req1.onreadystatechange = function() {
							if (this.readyState === 4) {
								req1.onreadystatechange = null;
								if (this.status === 200) {
									var resultdata = JSON.parse(this.response);
									if((resultdata.EntitySetName!=null)&&(resultdata.EntitySetName!="")&&(resultdata.EntitySetName!="undefined") )
								{
									  var collectedGuidElement =document.getElementById("hd_in_" + columnItem.name)as HTMLInputElement;
									
									  if((collectedGuidElement.value !="undefined")&&(collectedGuidElement.value !=null)&&(collectedGuidElement.value !=""))
									  {
										var collectedGuid:string=collectedGuidElement.value;
										collectedGuid=collectedGuid.replace("{","");
										collectedGuid=collectedGuid.replace("}","");

										data[schemaName+"@odata.bind"] = "/"+resultdata.EntitySetName+"(" +collectedGuid + ")";
									  }
								
								  
								}
								} 
								else {
									alert("smartgrid:03: and error occured while getting entity metadata for setting lookups"+this.statusText);
								}
							}
						};
						req1.send();
					
					
					
					//
					
					
					}
						} 
						else {
							alert("smartgrid:02: and error occured while getting entity metadata for setting lookups"+this.statusText);
						}
					}
				};
				req.send();

		
				field.value = "";
			  }
			
				else{
					var fieldOthers = document.getElementById("in_" + columnItem.name)as HTMLInputElement;
					if (fieldOthers.value != null && fieldOthers.value != "" && fieldOthers.value != "undefined") {

						data[columnItem.name] = fieldOthers.value;
					}
					fieldOthers.value = "";
					}
					
			}); 

			let test:string=this.contextObj.parameters.smartGridDataSet.getTargetEntityType().toString();
			// set the root relationship
	
			// @ts-ignore 
		let primaryLookupschemaName:string=	this.contextObj.parameters.primaryLookupName.raw.toString();
		// @ts-ignore 
		let entitySetName:string = this.contextObj.parameters.primaryEntitySet.raw.toString();
		
		// @ts-ignore 
	
			data[primaryLookupschemaName+"@odata.bind"] = "/"+entitySetName+"(" + this.contextObj.page.entityId + ")"; 
		
			
			let recordToCreate:string=this.contextObj.parameters.smartGridDataSet.getTargetEntityType().toString();
			let targetEntityCollectionName:string="";
			//get target entityset name
			var reqTargetCollection = new XMLHttpRequest();
						
			reqTargetCollection.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.0/EntityDefinitions(LogicalName='"+recordToCreate+"')", false);
			reqTargetCollection.setRequestHeader("OData-MaxVersion", "4.0");
			reqTargetCollection.setRequestHeader("OData-Version", "4.0");
			reqTargetCollection.setRequestHeader("Accept", "application/json");
			reqTargetCollection.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			reqTargetCollection.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
			reqTargetCollection.onreadystatechange = function() {
			if (this.readyState === 4) {
				reqTargetCollection.onreadystatechange = null;
				if (this.status === 200) {
					var resultdataTarget = JSON.parse(this.response);
					if((resultdataTarget.EntitySetName!=null)&&(resultdataTarget.EntitySetName!="")&&(resultdataTarget.EntitySetName!="undefined") )
				{
					
					targetEntityCollectionName=resultdataTarget.EntitySetName;

				  
				}
				} 
				else {
					alert("smartgrid:04: and error occured while getting entity metadata for target name"+this.statusText);
				}
			}
		};
		reqTargetCollection.send();

			//
		
			var reqPost = new XMLHttpRequest();
			reqPost.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/"+targetEntityCollectionName, false);
			reqPost.setRequestHeader("OData-MaxVersion", "4.0");
			reqPost.setRequestHeader("OData-Version", "4.0");
			reqPost.setRequestHeader("Accept", "application/json");
			reqPost.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			reqPost.onreadystatechange = function() {
						if (this.readyState === 4) {
							reqPost.onreadystatechange = null;
							if (this.status === 204) {
								var uri = this.getResponseHeader("OData-EntityId");
								//alert("New Item Added");
								thisRef.contextObj.navigation.openAlertDialog({text:"New "+thisRef.contextObj.parameters.smartGridDataSet.getTargetEntityType()+" Added", confirmButtonLabel : "Okay",})
								thisRef.contextObj.parameters.smartGridDataSet.refresh();
							} else {
								alert("An error occurred :("+this.statusText);
								thisRef.contextObj.parameters.smartGridDataSet.refresh();
							}
						}
					};
					reqPost.send(JSON.stringify(data));
			/* 		commented standard method as this is not working for all cases.
			this.contextObj.webAPI.createRecord(recordToCreate, data).then(function (response: ComponentFramework.EntityReference) {
				let id: string = response.id.toString();
			  alert("New Item Added");
			  thisRef.contextObj.parameters.smartGridDataSet.refresh();
			}, function (errorResponse: any) {
			  alert(errorResponse);
			  thisRef.contextObj.parameters.smartGridDataSet.refresh();	
			});  */
		 
		}

		  private openLookupDialogue(fieldID:string):void
		  {
			let targetName:string;
			  // get target entity
			  var thisRef = this;
			  var req = new XMLHttpRequest();
			  req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.0/EntityDefinitions(LogicalName='"+this.contextObj.parameters.smartGridDataSet.getTargetEntityType()+"')/Attributes(LogicalName='"+fieldID.slice(3)+"')", false);
			  req.setRequestHeader("OData-MaxVersion", "4.0");
			  req.setRequestHeader("OData-Version", "4.0");
			  req.setRequestHeader("Accept", "application/json");
			  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			  req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
			  req.onreadystatechange = function() {
				  if (this.readyState === 4) {
					  req.onreadystatechange = null;
					  if (this.status === 200) {
						  var result = JSON.parse(this.response);
						  if((result.Targets[0]!==null)&&(result.Targets[0]!=="")&&(result.Targets[0]!=="undefined") )
					  {

							targetName= result.Targets[0];
							var lookUpOptions: any = {
								entityTypes: [targetName]
							};


						    var lookUpPromise: any = thisRef.contextObj.utils.lookupObjects(lookUpOptions);
							lookUpPromise.then(
								// Callback method - invoked after user has selected an item from the lookup dialog
								// Data parameter is the item selected in the lookup dialog
								(data: ComponentFramework.EntityReference[]) => {
								if (data && data[0]) {
									// Get the ID and entityType of the record selected by the lookup
									let id:any = data[0].id;
									let entityType:string = data[0].etn!;
									// set the text box with details
									var currentField=document.getElementById(fieldID) as HTMLInputElement;
									currentField.value=data[0].name!;
									// set lookup details to  hidden fields
									var lookupHiddenValue=document.getElementById("hd_"+fieldID) as HTMLInputElement;

									lookupHiddenValue.value=id;
								}
								},
								(error: any) => {
								// Error handling code here
								
								}
							);
					  }
					  } 
					  else {
						  alert("smartgrid:01: and error occured while getting entity metadata"+this.statusText);
					  }
				  }
			  };
			  req.send();


			  //
			

			
				
		  }

}