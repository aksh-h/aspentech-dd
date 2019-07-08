import * as WitService from "TFS/WorkItemTracking/Services";
import * as VSSUtilsCore from "VSS/Utils/Core";
import Q = require("q");

export class BaseMultiValueControl {
    /**
     * Field name input for the control
     */
    public fieldName: string;
    public statefieldName: string;
    public CityfieldName: string;
    /**
     * The container to hold the control
     */
    protected containerElement: JQuery;
    protected StatecontainerElement: JQuery;
    protected CitycontainerElement: JQuery;
    /**
     * The container for error message display
     */
    private _errorPane: JQuery;
    private _StateerrorPane: JQuery;
    private _CityerrorPane: JQuery;

    private _flushing: boolean;
    private _bodyElement: HTMLBodyElement;

    /* Inherits from initalConfig to control if always show field border
     */
    private _showFieldBorder: boolean;
    /**
     * Store the last recorded window width to know
     * when we have been shrunk and should resize
     */
    private _windowWidth: number;
    private _minWindowWidthDelta: number = 10; // Minum change in window width to react to
    private _windowResizeThrottleDelegate: Function;

    constructor() {
        let initialConfig = VSS.getConfiguration();
        this._showFieldBorder = !!initialConfig.fieldBorder;

        this.containerElement = $(".container");
        if (this._showFieldBorder) {
            this.containerElement.addClass("fieldBorder");
        }
        this.StatecontainerElement = $(".statecontainer");
        if (this._showFieldBorder) {
            this.StatecontainerElement.addClass("fieldBorder");
        }
        this.CitycontainerElement = $(".citycontainer");
        if (this._showFieldBorder) {
            this.CitycontainerElement.addClass("fieldBorder");
        }
        this._errorPane = $("<div>").addClass("errorPane").appendTo(this.containerElement);
        this._StateerrorPane = $("<div>").addClass("errorPane").appendTo(this.StatecontainerElement);
        this._CityerrorPane = $("<div>").addClass("errorPane").appendTo(this.CitycontainerElement);
        var inputs: IDictionaryStringTo<string> = initialConfig.witInputs;

        this.fieldName = inputs["FieldName"];
        if (!this.fieldName) {
            this.showError("FieldName input has not been specified");
        }
        this.statefieldName = inputs["StateName"];
        if (!this.statefieldName) {
            this.showStateError("statefieldName input has not been specified");
        }
        this.CityfieldName = inputs["CityName"];
        if (!this.CityfieldName) {
            this.showStateError("CityfieldName input has not been specified");
        }
        this._windowResizeThrottleDelegate = VSSUtilsCore.throttledDelegate(this, 50, () => {
            this._windowWidth = window.innerWidth;
            this.resize();
        });

        this._windowWidth = window.innerWidth;
        $(window).resize(() => {
            if (Math.abs(this._windowWidth - window.innerWidth) > this._minWindowWidthDelta) {
                this._windowResizeThrottleDelegate.call(this);
            }
        });
    }

    /**
     * Initialize a new instance of Control
     */
    public initialize(): void {
        this.invalidate();
    }

    public Stateinitialize(): void {
        this.Stateinvalidate();
    }

    public Cityinitialize(): void {
        this.Cityinvalidate();
    }
    /**
     * Invalidate the control's value
     */
    public invalidate(): void {
        if (!this._flushing) {
            this._getCurrentFieldValue().then(
                (value: string) => {
                    this.setValue(value);
                }
            );
        }
        this.resize();
    }

    public clear(): void {
    }
    /**
     * Flushes the control's value to the field
     */
    protected flush(): void {
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.fieldName, this.getValue()).then(
                    (values) => {
                        this._flushing = false;
                    },
                    () => {
                        this._flushing = false;
                        this.showError("Error storing the field value");
                    }
                )
            }
        );
    }
    protected stateflush(): void {
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.statefieldName, this.getStateValue()).then(
                    (values) => {
                        this._flushing = false;
                    },
                    () => {
                        this._flushing = false;
                        this.showStateError("Error storing the field value");
                    }
                )
            }
        );
    }
    protected Cityflush(): void {
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.CityfieldName, this.getCityValue()).then(
                    (values) => {
                        this._flushing = false;
                    },
                    () => {
                        this._flushing = false;
                        this.showCityError("Error storing the field value");
                    }
                )
            }
        );
    }
    protected getValue(): string {
        return "";
    }

    protected ClearCity(): void {
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.CityfieldName, "").then(
                    (values) => {
                    },
                    () => {
                        this.showStateError("Error clearing the field value");
                    }
                )
            }
        );
    }
    protected ClearState(): void {
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.statefieldName,"").then(
                    (values) => {
                    },
                    () => {
                        this.showCityError("Error clearing the field value");
                    }
                )
            }
        );
    }


    protected setValue(value: string): void {
    }

    protected showError(error: string): void {
        this._errorPane.text(error);
        this._errorPane.show();
    }

    protected clearError() {
        this._errorPane.text("");
        this._errorPane.hide();
    }

    private _getCurrentFieldValue(): IPromise<string> {
        var defer = Q.defer();
        WitService.WorkItemFormService.getService().then(
            (service) => {
                service.getFieldValues([this.fieldName]).then(
                    (values) => {
                        defer.resolve(values[this.fieldName]);
                    },
                    () => {
                        this.showError("Error loading values for field: " + this.fieldName)
                    }
                );
            }
        );
        return defer.promise.then();
    }
    //#region  State
    public Stateinvalidate(): void {
        if (!this._flushing) {
            this._getCurrentStateFieldValue().then(
                (value: string) => {
                    this.setStateValue(value);
                }
            );
        }
        this.resize();
    }
    protected getStateValue(): string {
        return "";
    }

    protected setStateValue(value: string): void {
    }

    protected showStateError(error: string): void {
        this._StateerrorPane.text(error);
        this._StateerrorPane.show();
    }

    protected clearStateError() {
        this._StateerrorPane.text("");
        this._StateerrorPane.hide();
    }

    private _getCurrentStateFieldValue(): IPromise<string> {
        var defer = Q.defer();
        WitService.WorkItemFormService.getService().then(
            (service) => {
                service.getFieldValues([this.statefieldName]).then(
                    (values) => {
                        console.log(this.statefieldName);
                        defer.resolve(values[this.statefieldName]);
                        console.log(values[this.statefieldName]);
                    },
                    () => {
                        this.showStateError("Error loading values for field: " + this.statefieldName);
                    }
                );
            }
        );
        return defer.promise.then();
    }

    //#endregion State

    //#region City
    public Cityinvalidate(): void {
        if (!this._flushing) {
            this._getCurrentCityFieldValue().then(
                (value: string) => {
                    this.setCityValue(value);
                }
            );
        }
        this.resize();
    }
    protected getCityValue(): string {
        return "";
    }

    protected setCityValue(value: string): void {
    }

    protected showCityError(error: string): void {
        this._CityerrorPane.text(error);
        this._CityerrorPane.show();
    }

    protected clearCityError() {
        this._CityerrorPane.text("");
        this._CityerrorPane.hide();
    }

    private _getCurrentCityFieldValue(): IPromise<string> {
        var defer = Q.defer();
        WitService.WorkItemFormService.getService().then(
            (service) => {
                service.getFieldValues([this.CityfieldName]).then(
                    (values) => {
                        console.log(this.CityfieldName);
                        defer.resolve(values[this.CityfieldName]);
                        console.log(values[this.CityfieldName]);
                    },
                    () => {
                        this.showCityError("Error loading values for field: " + this.CityfieldName);
                    }
                );
            }
        );
        return defer.promise.then();
    }

    //#endregion
   
    protected resize() {
        this._bodyElement = <HTMLBodyElement>document.getElementsByTagName("body").item(0);
        // Cast as any until declarations are updated
        VSS.resize(null, this._bodyElement.offsetHeight);
    }
}