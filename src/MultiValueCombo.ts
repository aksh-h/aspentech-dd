import Q = require("q");
import * as VSSUtilsCore from "VSS/Utils/Core";
import { BaseMultiValueControl } from "./BaseMultiValueControl";
import { callDevApi } from "./RestCall";
import * as WitService from "TFS/WorkItemTracking/Services";

export class MultiValueCombo extends BaseMultiValueControl {
    /*
    * UI elements for the control.
    */
    public productName: string;
    private _selectedValuesWrapper: JQuery;
    private _selectedValuesContainer: JQuery;
    private _checkboxValuesContainer: JQuery;

    private _chevron: JQuery;

    private _suggestedValues: string[];

    private _valueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _valueToLabelMap: IDictionaryStringTo<JQuery>;

    private _StateselectedValuesWrapper: JQuery;
    private _StateselectedValuesContainer: JQuery;
    private _StatecheckboxValuesContainer: JQuery;

    private _Statechevron: JQuery;

    private _StatesuggestedValues: string[];

    private _StatevalueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _StatevalueToLabelMap: IDictionaryStringTo<JQuery>;

    private _CityselectedValuesWrapper: JQuery;
    private _CityselectedValuesContainer: JQuery;
    private _CitycheckboxValuesContainer: JQuery;

    private _Citychevron: JQuery;

    private _CitysuggestedValues: string[];

    private _CityvalueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _CityvalueToLabelMap: IDictionaryStringTo<JQuery>;

    private _maxSelectedToShow = 100;
    private _chevronDownClass = "bowtie-chevron-down-light";
    private _chevronUpClass = "bowtie-chevron-up-light";

    private _toggleThrottleDelegate: Function;
    private _StatetoggleThrottleDelegate: Function;
    private _CitytoggleThrottleDelegate: Function;
    /**
     * Initialize a new instance of MultiValueControl
     */
    public initialize(): void {
        this._selectedValuesWrapper = $("<div>").addClass("selectedValuesWrapper").appendTo(this.containerElement);
        this._selectedValuesContainer = $("<div>").addClass("selectedValuesContainer").attr("tabindex", "-1").appendTo(this._selectedValuesWrapper);
        this._chevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._selectedValuesWrapper);
        this._checkboxValuesContainer = $("<div>").addClass("checkboxValuesContainer").appendTo(this.containerElement);

        this._valueToCheckboxMap = {};
        this._valueToLabelMap = {};

        this._StateselectedValuesWrapper = $("<div>").addClass("StateselectedValuesWrapper").appendTo(this.StatecontainerElement);
        this._StateselectedValuesContainer = $("<div>").addClass("StateselectedValuesContainer").attr("tabindex", "-1").appendTo(this._StateselectedValuesWrapper);
        this._Statechevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._StateselectedValuesWrapper);
        this._StatecheckboxValuesContainer = $("<div>").addClass("StatecheckboxValuesContainer").appendTo(this.StatecontainerElement);

        this._StatevalueToCheckboxMap = {};
        this._StatevalueToLabelMap = {};

        this._CityselectedValuesWrapper = $("<div>").addClass("CityselectedValuesWrapper").appendTo(this.CitycontainerElement);
        this._CityselectedValuesContainer = $("<div>").addClass("CityselectedValuesContainer").attr("tabindex", "-1").appendTo(this._CityselectedValuesWrapper);
        this._Citychevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._CityselectedValuesWrapper);
        this._CitycheckboxValuesContainer = $("<div>").addClass("CitycheckboxValuesContainer").appendTo(this.CitycontainerElement);

        this._CityvalueToCheckboxMap = {};
        this._CityvalueToLabelMap = {};

        this._getSuggestedValues().then(
            (values: string[]) => {
                this._suggestedValues = values;
                this._populateCheckBoxes();
                super.initialize();
            }
        );
        this._getStateSuggestedValues().then(
            (values: string[]) => {
                this._StatesuggestedValues = values;
                this._populateStateCheckBoxes();
                super.Stateinitialize();
            }
        );

        this._getCitySuggestedValues().then(
            (values: string[]) => {
                this._CitysuggestedValues = values;
                this._populateCityCheckBoxes();
                super.Cityinitialize();
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

        // ---STATE--//

        this._StatetoggleThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 100, () => {
            this._toggleStateCheckBoxContainer();
        });

        this._StateselectedValuesWrapper.click(() => {
            this._StatetoggleThrottleDelegate.call(this);
            return false;
        });

        this._Statechevron.click(() => {
            this._StatetoggleThrottleDelegate.call(this);
            return false;
        });

        // --CITY-- //

        this._CitytoggleThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 100, () => {
            this._toggleCityCheckBoxContainer();
        });

        this._CityselectedValuesWrapper.click(() => {
            this._CitytoggleThrottleDelegate.call(this);
            return false;
        });

        this._Citychevron.click(() => {
            this._CitytoggleThrottleDelegate.call(this);
            return false;
        });
    }

    //#region Country

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

    private _showStateValues(values: string[]) {
        if (values.length <= 0) {
            this._StateselectedValuesContainer.append("<div class='noStateSelection'>No selection made</div>");
        } else {
            $.each(values, (i, value) => {
                var control;
                // only show first N selections and the rest as more.
                if (i < this._maxSelectedToShow) {
                    control = this._createSelectedStateValueControl(value);
                } else {
                    control = this._createSelectedStateValueControl(values.length - i + " more");
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
            // Add the select all method
            //let selectAllBox = this._createSelectAllControl();

            $.each(this._suggestedValues, (i, value) => {
                this._createCheckBoxControl(value);
            });
        }
    }

    private _createCheckBoxControl(value: string) {
        let label = this._createValueLabel(value);
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
        var url: string = inputs["Url"];
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }
    private _createValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "checkbox" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }
    // Convert unknown data type to string[]
    private _findArr(data: object): string[] {
        const inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var property: string = inputs.Property;
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
    private _createCheckBox(value: string, label: JQuery, action?: Function) {
        let checkbox = $("<input  />");
        checkbox.attr("type", "checkbox");
        checkbox.attr("name", value);
        checkbox.attr("value", value);
        checkbox.attr("tabindex", -1);
        checkbox.attr("id", "checkbox" + value);
        checkbox.attr("style", "visibility:hidden");
        checkbox.change((e) => {
            if (action) {
                action.call(this);
            }
            this.productName = value;
            this._getStateRespectiveValues(value).then(
                (values: string[]) => {
                    this._StatesuggestedValues = values;
                    this._populateOptions();
                }
            );
            this._refreshValue($(e.currentTarget).attr("value"));
            this.flush();
        });
        return checkbox;
    }
    //#endregion

    //#region State

    public GetStateSuggestedValue(): void {
        this._getStateSuggestedValues().then(
            (values: string[]) => {
                this._StatesuggestedValues = values;
                this._populateStateCheckBoxes();
                super.Stateinitialize();
            }
        );
    }

    public Stateclear(): void {
        var checkboxes: JQuery = $("input", this._StatecheckboxValuesContainer);
        var labels: JQuery = $(".checkboxLabel", this._StatecheckboxValuesContainer);
        checkboxes.prop("checked", false);
        checkboxes.removeClass("selectedCheckbox");
        this._StatecheckboxValuesContainer.empty();
    }
    protected getStateValue(): string {
        return this._StateselectedValuesContainer.text();
    }
    protected setStateValue(value: string): void {
        this.Stateclear();
        var selectedValues = value ? value.split(";") : [];

        this._showStateValues(selectedValues);
        $.each(selectedValues, (i, value) => {
            if (value) {
                // mark the checkbox as checked
                var checkbox = this._StatevalueToCheckboxMap[value];
                var label = this._StatevalueToLabelMap[value];
                if (checkbox) {
                    checkbox.prop("checked", true);
                    checkbox.addClass("selectedCheckbox");
                }
            }
        });
    }

    private _toggleStateCheckBoxContainer() {
        if (this._StatecheckboxValuesContainer.is(":visible")) {
            this._hideStateCheckBoxContainer();
        } else {
            this._showStateCheckBoxContainer();
        }
    }

    private _showStateCheckBoxContainer() {
        this._Statechevron.removeClass(this._chevronDownClass).addClass(this._chevronUpClass);
        this.StatecontainerElement.addClass("expanded").removeClass("collapsed");
        this._StatecheckboxValuesContainer.show();
        this.resize();
    }

    private _hideStateCheckBoxContainer() {
        this._Statechevron.removeClass(this._chevronUpClass).addClass(this._chevronDownClass);
        this.StatecontainerElement.removeClass("expanded").addClass("collapsed");
        this._StatecheckboxValuesContainer.hide();
        this.resize();
    }
    private _refreshStateValue(currentSelectedValue) {
        this._hideStateCheckBoxContainer();
        this._StateselectedValuesContainer.empty();
        var val = [currentSelectedValue];
        this._showStateValues(val);
    }

    private _createSelectedStateValueControl(value: string): JQuery {
        var control = $("<div />");
        if (value) {
            control.text(value);
            control.attr("title", value);
            // control.addClass("selected");
            this._StateselectedValuesContainer.empty();
            this._StateselectedValuesContainer.append(control);
        }

        return control;
    }
    /**
    * Populates the UI with the list of checkboxes to choose the value from.
    */
    private _populateStateCheckBoxes(): void {
        if (!this._StatesuggestedValues || this._StatesuggestedValues.length === 0) {
            this.showStateError("No values to select.");
        } else {
            this._StatecheckboxValuesContainer.empty();
            $.each(this._StatesuggestedValues, (i, value) => {
                this._createStateCheckBoxControl(value);
            });
        }
    }

    private _createStateCheckBoxControl(value: string) {
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var propValue = value;
        let label = this._createStateValueLabel(propValue);
        let checkbox = this._createStateCheckBox(propValue, label);
        let container = $("<div />").addClass("scheckboxContainer");
        checkbox.addClass("valueOption");
        this._StatevalueToCheckboxMap[value] = checkbox;
        this._StatevalueToLabelMap[value] = label;
        container.append(checkbox);
        container.append(label);
        this._StatecheckboxValuesContainer.append(container);
    }

    // loads values at first time
    private _getStateSuggestedValues(): IPromise<string[]> {
        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = inputs["StateUrl"];
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findStateArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }

    private _createStateValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "checkbox1" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }
    private _createStateCheckBox(value: string, label: JQuery, action?: Function) {
        let checkbox = $("<input  />");
        checkbox.attr("type", "checkbox");
        checkbox.attr("name", value);
        checkbox.attr("value", value);
        checkbox.attr("tabindex", 1);
        checkbox.attr("id", "checkbox1" + value);
        checkbox.attr("style", "visibility:hidden");
        checkbox.change((f) => {
            if (action) {
                action.call(this);
            }
            this.Cityclear();
            console.log("Product " + this.productName);
            console.log("area " + value);
            this._getCityRespectiveValues(value, this.productName).then(
                (values: string[]) => {
                    this._CitysuggestedValues = values;
                    if(this._CitysuggestedValues.length > 0) {
                        this._populateCityOptions();
                    }
                }
            );
            this._refreshStateValue($(f.currentTarget).attr("value"));
            this.stateflush();
        });
        return checkbox;
    }

    public _getStateRespectiveValues(value: string): IPromise<string[]> {
        // Removeing field values on the product chang
        $('.StateselectedValuesContainer').empty();
        this._StateselectedValuesContainer.append("<div class='noStateSelection'>No selection made</div>");
        this.ClearState();
        //this.stateflush();

        $('.CityselectedValuesContainer').empty();
        this._CityselectedValuesContainer.append("<div class='noCitySelection'>No selection made</div>");
        this.ClearCity();
        // this.Cityflush();

        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = "https://cascadingdd.azurewebsites.net/api/Area/Aspen Capital Cost Estimator";
        var urlSplit = url.split("/");
        urlSplit[urlSplit.length - 1] = value;
        url = urlSplit.join("/");
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findStateArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }

    private _populateOptions(): void {
        if (this._StatesuggestedValues.length >= 0) {
            this._StatecheckboxValuesContainer.empty();
            $.each(this._StatesuggestedValues, (i, value) => {
                this._createOptionControl(value);
            });
        }
    }
    private _createOptionControl(value: string) {
        let label = this._createStateValueLabel(value);
        let checkbox = this._createStateCheckBox(value, label);
        let container = $("<div />").addClass("scheckboxContainer");
        checkbox.addClass("valueOption");
        container.append(checkbox);
        container.append(label);
        this._StatecheckboxValuesContainer.append(container);
        this.StatecontainerElement.append(this._StatecheckboxValuesContainer);
    }

    private _findStateArr(data: object): string[] {
        const inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var property: string = inputs["StateProperty"];
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
    //#endregion state

    //#region City

    private _showCityValues(values: string[]) {
        if (values.length <= 0) {
            this._CityselectedValuesContainer.append("<div class='noCitySelection'>No selection made</div>");
        } else {
            $.each(values, (i, value) => {
                var control;
                // only show first N selections and the rest as more.
                if (i < this._maxSelectedToShow) {
                    control = this._createSelectedCityValueControl(value);
                } else {
                    control = this._createSelectedCityValueControl(values.length - i + " more");
                    control.attr("title", values.slice(i).join(";"));
                    return false;
                }
            });
        }
        this.resize();
    }

    public GetCitySuggestedValue(): void {
        this._getCitySuggestedValues().then(
            (values: string[]) => {
                this._CitysuggestedValues = values;
                this._populateCityCheckBoxes();
                super.Cityinitialize();
            }
        );
    }

    public Cityclear(): void {
        var checkboxes: JQuery = $("input", this._CitycheckboxValuesContainer);
        var labels: JQuery = $(".checkboxLabel", this._CitycheckboxValuesContainer);
        checkboxes.prop("checked", false);
        checkboxes.removeClass("selectedCheckbox");
        this._CitycheckboxValuesContainer.empty();
    }
    protected getCityValue(): string {
        return this._CityselectedValuesContainer.text();
    }
    protected setCityValue(value: string): void {
        this.Cityclear();
        var selectedValues = value ? value.split(";") : [];
        this._showCityValues(selectedValues);
        $.each(selectedValues, (i, value) => {
            if (value) {
                // mark the checkbox as checked
                var checkbox = this._CityvalueToCheckboxMap[value];
                var label = this._CityvalueToLabelMap[value];
                if (checkbox) {
                    checkbox.prop("checked", true);
                    checkbox.addClass("selectedCheckbox");
                }
            }
        });
    }

    private _toggleCityCheckBoxContainer() {
        if (this._CitycheckboxValuesContainer.is(":visible")) {
            this._hideCityCheckBoxContainer();
        } else {
            this._showCityCheckBoxContainer();
        }
    }

    private _showCityCheckBoxContainer() {
        this._Citychevron.removeClass(this._chevronDownClass).addClass(this._chevronUpClass);
        this.CitycontainerElement.addClass("expanded").removeClass("collapsed");
        this._CitycheckboxValuesContainer.show();
        this.resize();
    }

    private _hideCityCheckBoxContainer() {
        this._Citychevron.removeClass(this._chevronUpClass).addClass(this._chevronDownClass);
        this.CitycontainerElement.removeClass("expanded").addClass("collapsed");
        this._CitycheckboxValuesContainer.hide();
        this.resize();
    }
    private _refreshCityValue(currentCitySelectedValue) {
        this._hideCityCheckBoxContainer();
        this._CityselectedValuesContainer.empty();
        var val = [currentCitySelectedValue];
        this._showCityValues(val);
    }


    private _createSelectedCityValueControl(value: string): JQuery {
        var control = $("<div />");
        if (value) {
            control.text(value);
            control.attr("title", value);
            // control.addClass("selected");
            this._CityselectedValuesContainer.empty();
            this._CityselectedValuesContainer.append(control);
        }
        return control;
    }
    /**
    * Populates the UI with the list of checkboxes to choose the value from.
    */
    private _populateCityCheckBoxes(): void {
        if (this._CitysuggestedValues.length >= 0) {
            this._CitycheckboxValuesContainer.empty();
            $.each(this._CitysuggestedValues, (i, value) => {
                this._createCityCheckBoxControl(value);
            });
        }
    }

    private _createCityCheckBoxControl(value: string) {
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var propValue = value;
        let label = this._createCityValueLabel(propValue);
        let checkbox = this._createCityCheckBox(propValue, label);
        let container = $("<div />").addClass("ccheckboxContainer");
        checkbox.addClass("valueOption");
        this._CityvalueToCheckboxMap[value] = checkbox;
        this._CityvalueToLabelMap[value] = label;
        container.append(checkbox);
        container.append(label);
        this._CitycheckboxValuesContainer.append(container);
    }

    // loads values at first time
    private _getCitySuggestedValues(): IPromise<string[]> {
        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = inputs["CityUrl"];
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findCityArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }
    // ---End City---//

    private _createCityValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "checkbox2" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }

    private _createCityCheckBox(value: string, label: JQuery, action?: Function) {
        let checkbox = $("<input  />");
        checkbox.attr("type", "checkbox");
        checkbox.attr("name", value);
        checkbox.attr("value", value);
        checkbox.attr("tabindex", 2);
        checkbox.attr("id", "checkbox2" + value);
        checkbox.attr("style", "visibility:hidden");
        checkbox.change((g) => {
            if (action) {
                action.call(this);
            }
            this._refreshCityValue($(g.currentTarget).attr("value"));
            this.Cityflush();
        });
        return checkbox;
    }

    public _getCityRespectiveValues(value: string, productName: string): IPromise<string[]> {
        $('.CityselectedValuesContainer').empty();
        this._CityselectedValuesContainer.append("<div class='noCitySelection'>No selection made</div>");
        this.ClearCity();
        this._CitycheckboxValuesContainer.empty();

        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = "https://cascadingdd.azurewebsites.net/api/SubArea/" + productName + "/" + value;
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findCityArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }
    private _populateCityOptions(): void {
        if (this._CitysuggestedValues.length > 0) {
            this._CitycheckboxValuesContainer.empty();
            $.each(this._CitysuggestedValues, (i, value) => {
                this._CreateCityOptionControl(value);
            });
        }
    }
    private _CreateCityOptionControl(value: string) {
        let label = this._createCityValueLabel(value);
        let checkbox = this._createCityCheckBox(value, label);
        let container = $("<div />").addClass("ccheckboxContainer");
        checkbox.addClass("valueOption");
        container.append(checkbox);
        container.append(label);
        this._CitycheckboxValuesContainer.append(container);
        this.CitycontainerElement.append(this._CitycheckboxValuesContainer);
    }

    // Convert unknown data type to string[]
    private _findCityArr(data: object): string[] {
        const inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var property: string = inputs["CityProperty"];
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
    //#endregion
}