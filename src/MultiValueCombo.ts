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

    private _AreaselectedValuesWrapper: JQuery;
    private _AreaselectedValuesContainer: JQuery;
    private _AreacheckboxValuesContainer: JQuery;

    private _Areachevron: JQuery;

    private _AreasuggestedValues: string[];

    private _AreavalueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _AreavalueToLabelMap: IDictionaryStringTo<JQuery>;

    private _SubareaselectedValuesWrapper: JQuery;
    private _SubareaselectedValuesContainer: JQuery;
    private _SubareacheckboxValuesContainer: JQuery;

    private _Subareachevron: JQuery;

    private _SubareasuggestedValues: string[];

    private _SubareavalueToCheckboxMap: IDictionaryStringTo<JQuery>;
    private _SubareavalueToLabelMap: IDictionaryStringTo<JQuery>;

    private _maxSelectedToShow = 100;
    private _chevronDownClass = "bowtie-chevron-down-light";
    private _chevronUpClass = "bowtie-chevron-up-light";

    private _toggleThrottleDelegate: Function;
    private _AreatoggleThrottleDelegate: Function;
    private _SubareatoggleThrottleDelegate: Function;
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

        this._AreaselectedValuesWrapper = $("<div>").addClass("StateselectedValuesWrapper").appendTo(this.AreacontainerElement);
        this._AreaselectedValuesContainer = $("<div>").addClass("StateselectedValuesContainer").attr("tabindex", "-1").appendTo(this._AreaselectedValuesWrapper);
        this._Areachevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._AreaselectedValuesWrapper);
        this._AreacheckboxValuesContainer = $("<div>").addClass("StatecheckboxValuesContainer").appendTo(this.AreacontainerElement);

        this._AreavalueToCheckboxMap = {};
        this._AreavalueToLabelMap = {};

        this._SubareaselectedValuesWrapper = $("<div>").addClass("CityselectedValuesWrapper").appendTo(this.SubareacontainerElement);
        this._SubareaselectedValuesContainer = $("<div>").addClass("CityselectedValuesContainer").attr("tabindex", "-1").appendTo(this._SubareaselectedValuesWrapper);
        this._Subareachevron = $("<span />").addClass("bowtie-icon " + this._chevronDownClass).appendTo(this._SubareaselectedValuesWrapper);
        this._SubareacheckboxValuesContainer = $("<div>").addClass("CitycheckboxValuesContainer").appendTo(this.SubareacontainerElement);

        this._SubareavalueToCheckboxMap = {};
        this._SubareavalueToLabelMap = {};

        this._getSuggestedValues().then(
            (values: string[]) => {
                this._suggestedValues = values;
                this._populateCheckBoxes();
                super.initialize();
            }
        );
        this._getAreaSuggestedValues().then(
            (values: string[]) => {
                this._AreasuggestedValues = values;
                this._populateAreaCheckBoxes();
                super.Areainitialize();
            }
        );

        this._getSubareaSuggestedValues().then(
            (values: string[]) => {
                this._SubareasuggestedValues = values;
                this._populateSubareaCheckBoxes();
                super.Subareainitialize();
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

        // ---Area--//

        this._AreatoggleThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 100, () => {
            this._toggleAreaCheckBoxContainer();
        });

        this._AreaselectedValuesWrapper.click(() => {
            this._AreatoggleThrottleDelegate.call(this);
            return false;
        });

        this._Areachevron.click(() => {
            this._AreatoggleThrottleDelegate.call(this);
            return false;
        });

        // --Subarea-- //

        this._SubareatoggleThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 100, () => {
            this._toggleSubareaCheckBoxContainer();
        });

        this._SubareaselectedValuesWrapper.click(() => {
            this._SubareatoggleThrottleDelegate.call(this);
            return false;
        });

        this._Subareachevron.click(() => {
            this._SubareatoggleThrottleDelegate.call(this);
            return false;
        });
    }

    //#region Product

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

    private _showAreaValues(values: string[]) {
        if (values.length <= 0) {
            this._AreaselectedValuesContainer.append("<div class='noAreaSelection'>No selection made</div>");
        } else {
            $.each(values, (i, value) => {
                var control;
                // only show first N selections and the rest as more.
                if (i < this._maxSelectedToShow) {
                    control = this._createSelectedAreaValueControl(value);
                } else {
                    control = this._createSelectedAreaValueControl(values.length - i + " more");
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
        var url: string = inputs["ProductUrl"];
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
        var property: string = inputs["ProductProperty"];
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
            this._getAreaRespectiveValues(value).then(
                (values: string[]) => {
                    this._AreasuggestedValues = values;
                    this._populateOptions();
                }
            );
            this._refreshValue($(e.currentTarget).attr("value"));
            this.flush();
        });
        return checkbox;
    }
    //#endregion

    //#region Area

    public Areaclear(): void {
        var checkboxes: JQuery = $("input", this._AreacheckboxValuesContainer);
        var labels: JQuery = $(".checkboxLabel", this._AreacheckboxValuesContainer);
        checkboxes.prop("checked", false);
        checkboxes.removeClass("selectedCheckbox");
        this._AreacheckboxValuesContainer.empty();
    }
    protected getAreaValue(): string {
        return this._AreaselectedValuesContainer.text();
    }
    protected setAreaValue(value: string): void {
        this.Areaclear();
        var selectedValues = value ? value.split(";") : [];

        this._showAreaValues(selectedValues);
        $.each(selectedValues, (i, value) => {
            if (value) {
                // mark the checkbox as checked
                var checkbox = this._AreavalueToCheckboxMap[value];
                var label = this._AreavalueToLabelMap[value];
                if (checkbox) {
                    checkbox.prop("checked", true);
                    checkbox.addClass("selectedCheckbox");
                }
            }
        });
    }

    private _toggleAreaCheckBoxContainer() {
        if (this._AreacheckboxValuesContainer.is(":visible")) {
            this._hideAreaCheckBoxContainer();
        } else {
            this._showAreaCheckBoxContainer();
        }
    }

    private _showAreaCheckBoxContainer() {
        this._Areachevron.removeClass(this._chevronDownClass).addClass(this._chevronUpClass);
        this.AreacontainerElement.addClass("expanded").removeClass("collapsed");
        this._AreacheckboxValuesContainer.show();
        this.resize();
    }

    private _hideAreaCheckBoxContainer() {
        this._Areachevron.removeClass(this._chevronUpClass).addClass(this._chevronDownClass);
        this.AreacontainerElement.removeClass("expanded").addClass("collapsed");
        this._AreacheckboxValuesContainer.hide();
        this.resize();
    }
    private _refreshAreaValue(currentSelectedValue) {
        this._hideAreaCheckBoxContainer();
        this._AreaselectedValuesContainer.empty();
        var val = [currentSelectedValue];
        this._showAreaValues(val);
    }

    private _createSelectedAreaValueControl(value: string): JQuery {
        var control = $("<div />");
        if (value) {
            control.text(value);
            control.attr("title", value);
            // control.addClass("selected");
            this._AreaselectedValuesContainer.empty();
            this._AreaselectedValuesContainer.append(control);
        }

        return control;
    }
    /**
    * Populates the UI with the list of checkboxes to choose the value from.
    */
    private _populateAreaCheckBoxes(): void {
        if (this._AreasuggestedValues.length >= 0) {
            this._AreacheckboxValuesContainer.empty();
            $.each(this._AreasuggestedValues, (i, value) => {
                this._createAreaCheckBoxControl(value);
            });
        }
    }

    private _createAreaCheckBoxControl(value: string) {
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var propValue = value;
        let label = this._createAreaValueLabel(propValue);
        let checkbox = this._createAreaCheckBox(propValue, label);
        let container = $("<div />").addClass("scheckboxContainer");
        checkbox.addClass("valueOption");
        this._AreavalueToCheckboxMap[value] = checkbox;
        this._AreavalueToLabelMap[value] = label;
        container.append(checkbox);
        container.append(label);
        this._AreacheckboxValuesContainer.append(container);
    }

    // loads values at first time
    private _getAreaSuggestedValues(): IPromise<string[]> {
        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = inputs["AreaUrl"];
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findAreaArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }

    private _createAreaValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "checkbox1" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }
    private _createAreaCheckBox(value: string, label: JQuery, action?: Function) {
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
            this.Subareaclear();
            this._getSubareaRespectiveValues(value, this.productName).then(
                (values: string[]) => {
                    this._SubareasuggestedValues = values;
                    if (this._SubareasuggestedValues.length > 0) {
                        this._populateSubareaOptions();
                    }
                }
            );
            this._refreshAreaValue($(f.currentTarget).attr("value"));
            this.Areaflush();
        });
        return checkbox;
    }

    public _getAreaRespectiveValues(value: string): IPromise<string[]> {
        // Removeing field values on the product chang
        $('.StateselectedValuesContainer').empty();
        this._AreaselectedValuesContainer.append("<div class='noAreaSelection'>No selection made</div>");
        this.ClearArea();

        $('.CityselectedValuesContainer').empty();
        this._SubareaselectedValuesContainer.append("<div class='noSubareaSelection'>No selection made</div>");
        this.ClearSubarea();
        // this.Subareaflush();

        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        console.log(inputs["AreaUrl"]);
        var url: string = inputs["AreaUrl"]+"/"+value;
        console.log(url);
        // var urlSplit = url.split("/");
        // urlSplit[urlSplit.length - 1] = value;
        // url = urlSplit.join("/");
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findAreaArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }

    private _populateOptions(): void {
        if (this._AreasuggestedValues.length >= 0) {
            this._AreacheckboxValuesContainer.empty();
            $.each(this._AreasuggestedValues, (i, value) => {
                this._createOptionControl(value);
            });
        }
    }
    private _createOptionControl(value: string) {
        let label = this._createAreaValueLabel(value);
        let checkbox = this._createAreaCheckBox(value, label);
        let container = $("<div />").addClass("scheckboxContainer");
        checkbox.addClass("valueOption");
        container.append(checkbox);
        container.append(label);
        this._AreacheckboxValuesContainer.append(container);
        this.AreacontainerElement.append(this._AreacheckboxValuesContainer);
    }

    private _findAreaArr(data: object): string[] {
        const inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var property: string = inputs["AreaProperty"];
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
    //#endregion Area

    //#region Subarea

    private _showSubareaValues(values: string[]) {
        if (values.length <= 0) {
            this._SubareaselectedValuesContainer.append("<div class='noSubareaSelection'>No selection made</div>");
        } else {
            $.each(values, (i, value) => {
                var control;
                // only show first N selections and the rest as more.
                if (i < this._maxSelectedToShow) {
                    control = this._createSelectedSubareaValueControl(value);
                } else {
                    control = this._createSelectedSubareaValueControl(values.length - i + " more");
                    control.attr("title", values.slice(i).join(";"));
                    return false;
                }
            });
        }
        this.resize();
    }

    public Subareaclear(): void {
        var checkboxes: JQuery = $("input", this._SubareacheckboxValuesContainer);
        var labels: JQuery = $(".checkboxLabel", this._SubareacheckboxValuesContainer);
        checkboxes.prop("checked", false);
        checkboxes.removeClass("selectedCheckbox");
        this._SubareacheckboxValuesContainer.empty();
    }
    protected getSubareaValue(): string {
        return this._SubareaselectedValuesContainer.text();
    }
    protected setSubareaValue(value: string): void {
        this.Subareaclear();
        var selectedValues = value ? value.split(";") : [];
        this._showSubareaValues(selectedValues);
        $.each(selectedValues, (i, value) => {
            if (value) {
                // mark the checkbox as checked
                var checkbox = this._SubareavalueToCheckboxMap[value];
                var label = this._SubareavalueToLabelMap[value];
                if (checkbox) {
                    checkbox.prop("checked", true);
                    checkbox.addClass("selectedCheckbox");
                }
            }
        });
    }

    private _toggleSubareaCheckBoxContainer() {
        if (this._SubareacheckboxValuesContainer.is(":visible")) {
            this._hideSubareaCheckBoxContainer();
        } else {
            this._showSubareaCheckBoxContainer();
        }
    }

    private _showSubareaCheckBoxContainer() {
        this._Subareachevron.removeClass(this._chevronDownClass).addClass(this._chevronUpClass);
        this.SubareacontainerElement.addClass("expanded").removeClass("collapsed");
        this._SubareacheckboxValuesContainer.show();
        this.resize();
    }

    private _hideSubareaCheckBoxContainer() {
        this._Subareachevron.removeClass(this._chevronUpClass).addClass(this._chevronDownClass);
        this.SubareacontainerElement.removeClass("expanded").addClass("collapsed");
        this._SubareacheckboxValuesContainer.hide();
        this.resize();
    }
    private _refreshSubareaValue(currentSubareaSelectedValue) {
        this._hideSubareaCheckBoxContainer();
        this._SubareaselectedValuesContainer.empty();
        var val = [currentSubareaSelectedValue];
        this._showSubareaValues(val);
    }


    private _createSelectedSubareaValueControl(value: string): JQuery {
        var control = $("<div />");
        if (value) {
            control.text(value);
            control.attr("title", value);
            // control.addClass("selected");
            this._SubareaselectedValuesContainer.empty();
            this._SubareaselectedValuesContainer.append(control);
        }
        return control;
    }
    /**
    * Populates the UI with the list of checkboxes to choose the value from.
    */
    private _populateSubareaCheckBoxes(): void {
        if (this._SubareasuggestedValues.length >= 0) {
            this._SubareacheckboxValuesContainer.empty();
            $.each(this._SubareasuggestedValues, (i, value) => {
                this._createSubareaCheckBoxControl(value);
            });
        }
    }

    private _createSubareaCheckBoxControl(value: string) {
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var propValue = value;
        let label = this._createSubareaValueLabel(propValue);
        let checkbox = this._createSubareaCheckBox(propValue, label);
        let container = $("<div />").addClass("ccheckboxContainer");
        checkbox.addClass("valueOption");
        this._SubareavalueToCheckboxMap[value] = checkbox;
        this._SubareavalueToLabelMap[value] = label;
        container.append(checkbox);
        container.append(label);
        this._SubareacheckboxValuesContainer.append(container);
    }

    // loads values at first time
    private _getSubareaSuggestedValues(): IPromise<string[]> {
        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = inputs["SubareaUrl"];
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findSubareaArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }
    // ---End Subarea---//

    private _createSubareaValueLabel(value: string) {
        let label = $("<label />");
        label.attr("for", "checkbox2" + value);
        label.text(value);
        label.attr("title", value);
        label.addClass("checkboxLabel");
        return label;
    }

    private _createSubareaCheckBox(value: string, label: JQuery, action?: Function) {
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
            this._refreshSubareaValue($(g.currentTarget).attr("value"));
            this.Subareaflush();
        });
        return checkbox;
    }

    public _getSubareaRespectiveValues(value: string, productName: string): IPromise<string[]> {
        $('.CityselectedValuesContainer').empty();
        this._SubareaselectedValuesContainer.append("<div class='noSubareaSelection'>No selection made</div>");
        this.ClearSubarea();
        this._SubareacheckboxValuesContainer.empty();

        var defer = Q.defer<any>();
        var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var url: string = inputs["SubareaUrl"] + "/" + productName + "/" + value;
        callDevApi(url, "GET", undefined, undefined, (data) => {
            defer.resolve(this._findSubareaArr(data));
        }, (error) => {
            defer.reject(error);
        });
        return defer.promise;
    }
    private _populateSubareaOptions(): void {
        if (this._SubareasuggestedValues.length > 0) {
            this._SubareacheckboxValuesContainer.empty();
            $.each(this._SubareasuggestedValues, (i, value) => {
                this._CreateSubareaOptionControl(value);
            });
        }
    }
    private _CreateSubareaOptionControl(value: string) {
        let label = this._createSubareaValueLabel(value);
        let checkbox = this._createSubareaCheckBox(value, label);
        let container = $("<div />").addClass("ccheckboxContainer");
        checkbox.addClass("valueOption");
        container.append(checkbox);
        container.append(label);
        this._SubareacheckboxValuesContainer.append(container);
        this.SubareacontainerElement.append(this._SubareacheckboxValuesContainer);
    }

    // Convert unknown data type to string[]
    private _findSubareaArr(data: object): string[] {
        const inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
        var property: string = inputs["SubareaProperty"];
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