import $ = require('jquery');
import Q = require("q");
import * as VSSUtilsCore from "VSS/Utils/Core";
import { BaseMultiValueControl } from "./BaseMultiValueControl";
import { callDevApi } from "./RestCall";
import * as WitService from "TFS/WorkItemTracking/Services";

export class MultiValueCombo extends BaseMultiValueControl {
    /*
    * UI elements for the control.
    */
   
    private selectFamily:string;
    private selectProduct:string;

    private _maxSelectedToShow = 200;
    private _chevronDownClass = "bowtie-chevron-down-light";
    private _chevronUpClass = "bowtie-chevron-up-light";

    private _selectedValuesWrapper: JQuery;
    private _selectedValuesContainer: JQuery;
    private _checkboxValuesContainer: JQuery;
    private _chevron: JQuery;
    private _suggestedValues: string[];
    private _toggleThrottleDelegate: Function;

    private _valueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _valueToLabelMap: IDictionaryStringTo<JQuery>;


    private _ProductselectedValuesWrapper: JQuery;
    private _ProductselectedValuesContainer: JQuery;
    private _ProductcheckboxValuesContainer: JQuery;
    private _Productchevron: JQuery;
    private _ProductsuggestedValues: string[];
    private _ProducttoggleThrottleDelegate: Function;
    private _ProductvalueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _ProductvalueToLabelMap: IDictionaryStringTo<JQuery>;

    /**
     * Initialize a new instance of MultiValueControl
     */
    public initialize(): void {
        //#region Family
        this._selectedValuesWrapper = $("<div>").addClass("selectedValuesWrapper").appendTo(this.containerElement);
        this._selectedValuesContainer = $("<div>").addClass("selectedValuesContainer").attr("tabindex", "-1").appendTo(this._selectedValuesWrapper);
        this._chevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._selectedValuesWrapper);
        this._checkboxValuesContainer = $("<div>").addClass("checkboxValuesContainer").appendTo(this.containerElement);

        this._valueToCheckboxMap = {};
        this._valueToLabelMap = {};

        this._getSuggestedValues().then(
            (values: string[]) => {
                this._suggestedValues = values;
                this._populateCheckBoxes();
                super.initialize();
            }
        );
        this._toggleThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 100, () => {
            this._toggleCheckBoxContainer();
        });

        this._selectedValuesWrapper.click(() => {
            this._toggleThrottleDelegate.call(this);
            return false;
        });

        this._chevron.click(() => {
            this._toggleThrottleDelegate.call(this);
            return false;
        });

        //#endregion

        //#region Product
        this._ProductselectedValuesWrapper = $("<div>").addClass("productselectedValuesWrapper").appendTo(this.ProductcontainerElement);
        this._ProductselectedValuesContainer = $("<div>").addClass("productselectedValuesContainer").attr("tabindex", "-1").appendTo(this._ProductselectedValuesWrapper);
        this._Productchevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._ProductselectedValuesWrapper);
        this._ProductcheckboxValuesContainer = $("<div>").addClass("productcheckboxValuesContainer").appendTo(this.ProductcontainerElement);
        this._ProductvalueToCheckboxMap = {};
        this._ProductvalueToLabelMap = {};

        this._ProductgetSuggestedValues().then(
            (values: string[]) => {
                this._ProductsuggestedValues = values;
                this._ProductpopulateCheckBoxes();
                super.Productinitialize();
            }
        );
        this._ProducttoggleThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 100, () => {
            this._ProducttoggleCheckBoxContainer();
        });

        this._ProductselectedValuesWrapper.click(() => {
            this._ProducttoggleThrottleDelegate.call(this);
            return false;
        });

        this._Productchevron.click(() => {
            this._ProducttoggleThrottleDelegate.call(this);
            return false;
        });

        //#endregion
    }

    //#region Family

    public clear(): void {
        var checkboxes: JQuery = $("input", this._checkboxValuesContainer);
        var labels: JQuery = $(".checkboxLabel", this._checkboxValuesContainer);
        checkboxes.prop("checked", false);
        checkboxes.removeClass("selectedCheckbox");
        this._selectedValuesContainer.empty();
    }
    protected getValue(): string {
        return this._selectedValuesContainer.text();
    }

    protected setValue(value: string): void {
        this.clear();
        var selectedValues = value ? value.split(";") : [];

        this._showValues(selectedValues);
        $.each(selectedValues, (i, value) => {
            if (value) {
                // mark the checkbox as checked
                var checkbox = this._valueToCheckboxMap[value];
                var label = this._valueToLabelMap[value];
                if (checkbox) {
                    checkbox.prop("checked", true);
                    checkbox.addClass("selectedCheckbox");
                }
            }
        });
    }

    private _toggleCheckBoxContainer() {
        if (this._checkboxValuesContainer.is(":visible")) {
            this._hideCheckBoxContainer();
        } else {
            this._showCheckBoxContainer();
        }
    }

    private _showCheckBoxContainer() {
        this._chevron.removeClass(this._chevronDownClass).addClass(this._chevronUpClass);
        this.containerElement.addClass("expanded").removeClass("collapsed");
        this._checkboxValuesContainer.show();
        this.resize();
    }

    private _hideCheckBoxContainer() {
        this._chevron.removeClass(this._chevronUpClass).addClass(this._chevronDownClass);
        this.containerElement.removeClass("expanded").addClass("collapsed");
        this._checkboxValuesContainer.hide();
        this.resize();
    }

    private _showValues(values: string[]) {
        if (values.length <= 0) {
            this._selectedValuesContainer.append("<div class='noSelection'>No selection made</div>");
        } else {
            $.each(values, (i, value) => {
                var control;
                // only show first N selections and the rest as more.
                if (i < this._maxSelectedToShow) {
                    control = this._createSelectedValueControl(value);
                } else {
                    control = this._createSelectedValueControl(values.length - i + " more");
                    control.attr("title", values.slice(i).join(";"));
                    return false;
                }
            });
        }
        this.resize();
    }

    private _refreshValue(currentSelectedValue) {
        this._hideCheckBoxContainer();
        this._selectedValuesContainer.empty();
        var val = [currentSelectedValue];
        this._showValues(val);
    }

    private _createSelectedValueControl(value: string): JQuery {
        var control = $("<div />");
        if (value) {
            control.text(value);
            control.attr("title", value);
            // control.addClass("selected");
            this._selectedValuesContainer.empty();
            this._selectedValuesContainer.append(control);
        }
        return control;
    }
    /**
     * Populates the UI with the list of checkboxes to choose the value from.
     */
    private _populateCheckBoxes(): void {
        if (!this._suggestedValues || this._suggestedValues.length === 0) {
            this.showError("No values to select.");
        } else {
            $.each(this._suggestedValues, (i, value) => {
                this._createCheckBoxControl(value);
            });
        }
    }

    private _createCheckBoxControl(value: string) {
        let label = this._createFamilyValueLabel(value);
        let checkbox = this._createCheckBox(value, label);
        let container = $("<div />").addClass("checkboxContainer");
        checkbox.addClass("valueOption");
        this._valueToCheckboxMap[value] = checkbox;
        this._valueToLabelMap[value] = label;
        container.append(checkbox);
        container.append(label);
        this._checkboxValuesContainer.append(container);
    }

    private _getSuggestedValues(): IPromise<string[]> {
        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = inputs["FamilyUrl"];
        var property: string = inputs["FamilyProperty"];
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findArr(property,data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }
  
    private _createCheckBox(value: string, label: JQuery, action?: Function) {
        let checkbox = $("<input  />");
        checkbox.attr("type", "checkbox");
        checkbox.attr("name", value);
        checkbox.attr("value", value);
        checkbox.attr("tabindex", -1);
        checkbox.attr("id", "optionFamily" + value);
        checkbox.attr("style", "visibility:hidden");
        checkbox.change((e) => {
            if (action) {
                action.call(this);
            }
            this.selectFamily = value;
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var url: string = inputs["ProductUrl"];
            var property: string = inputs["ProductProperty"];
            this._getRespectiveValues(url,property,this.selectFamily).then(
                (values: string[]) => {
                    this._ProductsuggestedValues = values;
                    if (this._ProductsuggestedValues.length > 0){
                        this._ProductpopulateCheckBoxes();
                    }
                }
            );
            this._refreshValue($(e.currentTarget).attr("value"));
            this.flush();
        });
        return checkbox;
    }
    
    //#endregion

    //#region Product

        public Productclear(): void {
            var checkboxes: JQuery = $("input", this._ProductcheckboxValuesContainer);
            var labels: JQuery = $(".checkboxLabel", this._ProductcheckboxValuesContainer);
            checkboxes.prop("checked", false);
            checkboxes.removeClass("selectedCheckbox");
            this._ProductselectedValuesContainer.empty();
        }
        protected ProductgetValue(): string {
            return this._ProductselectedValuesContainer.text();
        }

        protected ProductsetValue(value: string): void {
            this.Productclear();
            var selectedValues = value ? value.split(";") : [];

            this._ProductshowValues(selectedValues);
            $.each(selectedValues, (i, value) => {
                if (value) {
                    // mark the checkbox as checked
                    var checkbox = this._ProductvalueToCheckboxMap[value];
                    var label = this._ProductvalueToLabelMap[value];
                    if (checkbox) {
                        checkbox.prop("checked", true);
                        checkbox.addClass("selectedCheckbox");
                    }
                }
            });
        }
        private _hideProductCheckBoxContainer() {
            this._Productchevron.removeClass(this._chevronUpClass).addClass(this._chevronDownClass);
            this.ProductcontainerElement.removeClass("expanded").addClass("collapsed");
            this._ProductcheckboxValuesContainer.hide();
            this.resize();
        }

        private _showProductCheckBoxContainer() {
            this._Productchevron.removeClass(this._chevronDownClass).addClass(this._chevronUpClass);
            this.ProductcontainerElement.addClass("expanded").removeClass("collapsed");
            this._ProductcheckboxValuesContainer.show();
            this.resize();
        }

        private _refreshProductValue(currentSelectedValue) {
            this._hideProductCheckBoxContainer();
            this._ProductselectedValuesContainer.empty();
            var val = [currentSelectedValue];
            this._ProductshowValues(val);
        }

        private _ProductcreateSelectedValueControl(value: string): JQuery {
            var control = $("<div />");
            if (value) {
                control.text(value);
                control.attr("title", value);
                // control.addClass("selected");
                this._ProductselectedValuesContainer.empty();
                this._ProductselectedValuesContainer.append(control);
            }
            return control;
        }

        private _ProductshowValues(values: string[]) {
            if (values.length <= 0) {
                this._ProductselectedValuesContainer.append("<div class='noProductSelection'>No selection made</div>");
            } else {
                $.each(values, (i, value) => {
                    var control;
                    // only show first N selections and the rest as more.
                    if (i < this._maxSelectedToShow) {
                        control = this._ProductcreateSelectedValueControl(value);
                    } else {
                        control = this._ProductcreateSelectedValueControl(values.length - i + " more");
                        control.attr("title", values.slice(i).join(";"));
                        return false;
                    }
                });
            }
            this.resize();
        }

        private _ProducttoggleCheckBoxContainer() {
            if (this._ProductcheckboxValuesContainer.is(":visible")) {
                this._hideProductCheckBoxContainer();
            } else {
                this._showProductCheckBoxContainer();
            }
        }
        private _ProductgetSuggestedValues(): IPromise<string[]> {
            var defer = Q.defer<any>();
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var url: string = inputs["ProductUrl"];
            var property: string = inputs["ProductProperty"];
            callDevApi(url, "GET", undefined, undefined, (data) => {
                defer.resolve(this._findArr(property,data));
            }, (error) => {
                defer.reject(error);
            });
            return defer.promise;
        }

        private _ProductpopulateCheckBoxes(): void {
            if (!this._ProductsuggestedValues || this._ProductsuggestedValues.length === 0) {
                this.showProductError("No values to select.");
            } else {
                $.each(this._ProductsuggestedValues, (i, value) => {
                    this._createProductCheckBoxControl(value);
                });
            }
        }
        private _createProductCheckBoxControl(value: string) {
            let label = this._createProductValueLabel(value);
            let checkbox = this._createProductCheckBox(value, label);
            let container = $("<div />").addClass("pcheckboxContainer");
            checkbox.addClass("valueOption");
            this._ProductvalueToCheckboxMap[value] = checkbox;
            this._ProductvalueToLabelMap[value] = label;
            container.append(checkbox);
            container.append(label);
            this._ProductcheckboxValuesContainer.empty().append(container);
        }
       
        private _createProductCheckBox(value: string, label: JQuery, action?: Function) {
            let checkbox = $("<input  />");
            checkbox.attr("type", "checkbox");
            checkbox.attr("name", value);
            checkbox.attr("value", value);
            checkbox.attr("tabindex", 1);
            checkbox.attr("id", "optionProduct" + value);
            checkbox.attr("style", "visibility:hidden");
            checkbox.change((f) => {
                if (action) {
                    action.call(this);
                }
                this._refreshProductValue($(f.currentTarget).attr("value"));
                this.Productflush();
            });
            return checkbox;
        }
       
    //#endregion

    // Convert unknown data type to string[]
    private _findArr(property: string, data: object): string[] {
        const inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        // Look for an array: object itself or one of its properties
        const objs: object[] = [data];
        for (let obj = objs.shift(); obj; obj = objs.shift()) {
            if (Array.isArray(obj)) {
                // If configuration has a the Property property set then map from objects to strings
                // Otherwise assume already strings
                return property ? obj.map(o => o[property]) : obj;
            } else if (typeof obj === "object") {
                for (const key in obj) {
                    objs.push(obj[key]);
                }
            }
        }
    }
  
    public _getRespectiveValues (api:string,property:string,value: string): IPromise<string[]> {
        var defer = Q.defer<any>();
        var url: string = api;
        var urlSplit = url.split("/");
        urlSplit[urlSplit.length - 1] = value;
        url = urlSplit.join("/");
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findArr(property,data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }

    private _createValueLabel(identity: string,value: string) {
        let label = $("<label />");
        label.attr("for", identity + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }

    private _createFamilyValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "optionFamily" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }
    private _createProductValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "optionProduct"+ value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }
    private _createAreaValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "optionArea" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }
}