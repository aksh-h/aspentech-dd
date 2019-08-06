import * as WitService from "TFS/WorkItemTracking/Services";
import * as VSSUtilsCore from "VSS/Utils/Core";
import Q = require("q");

export class BaseMultiValueControl {
    /**
     * Field name input for the control
     */
    public fieldName: string;
    public AreafieldName: string;
    public SubareafieldName: string;
    /**
     * The container to hold the control
     */
    protected containerElement: JQuery;
    protected AreacontainerElement: JQuery;
    protected SubareacontainerElement: JQuery;
    /**
     * The container for error message display
     */
    private _errorPane: JQuery;
    private _AreaerrorPane: JQuery;
    private _SubareaerrorPane: JQuery;

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
        this.AreacontainerElement = $(".statecontainer");
        if (this._showFieldBorder) {
            this.AreacontainerElement.addClass("fieldBorder");
        }
        this.SubareacontainerElement = $(".citycontainer");
        if (this._showFieldBorder) {
            this.SubareacontainerElement.addClass("fieldBorder");
        }
        this._errorPane = $("<div>").addClass("errorPane").appendTo(this.containerElement);
        this._AreaerrorPane = $("<div>").addClass("errorPane").appendTo(this.AreacontainerElement);
        this._SubareaerrorPane = $("<div>").addClass("errorPane").appendTo(this.SubareacontainerElement);
        var inputs: IDictionaryStringTo<string> = initialConfig.witInputs;

        this.fieldName = inputs["Product"];
        if (!this.fieldName) {
            this.showError("Product input has not been specified");
        }
        this.AreafieldName = inputs["AreaName"];
        if (!this.AreafieldName) {
            this.showAreaError("AreaName input has not been specified");
        }
        this.SubareafieldName = inputs["SubareaName"];
        if (!this.SubareafieldName) {
            this.showAreaError("SubareaName input has not been specified");
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

    public Areainitialize(): void {
        this.Areainvalidate();
    }

    public Subareainitialize(): void {
        this.Subareainvalidate();
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
    protected Areaflush(): void {
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.AreafieldName, this.getAreaValue()).then(
                    (values) => {
                        this._flushing = false;
                    },
                    () => {
                        this._flushing = false;
                        this.showAreaError("Error storing the field value");
                    }
                )
            }
        );
    }
    protected Subareaflush(): void {
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.SubareafieldName, this.getSubareaValue()).then(
                    (values) => {
                        this._flushing = false;
                    },
                    () => {
                        this._flushing = false;
                        this.showSubareaError("Error storing the field value");
                    }
                )
            }
        );
    }
    protected getValue(): string {
        return "";
    }

    protected ClearSubarea(): void {
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.SubareafieldName, "").then(
                    (values) => {
                    },
                    () => {
                        this.showAreaError("Error clearing the field value");
                    }
                )
            }
        );
    }
    protected ClearArea(): void {
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.AreafieldName,"").then(
                    (values) => {
                    },
                    () => {
                        this.showSubareaError("Error clearing the field value");
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
    //#region  Area
    public Areainvalidate(): void {
        if (!this._flushing) {
            this._getCurrentAreaFieldValue().then(
                (value: string) => {
                    this.setAreaValue(value);
                }
            );
        }
        this.resize();
    }
    protected getAreaValue(): string {
        return "";
    }

    protected setAreaValue(value: string): void {
    }

    protected showAreaError(error: string): void {
        this._AreaerrorPane.text(error);
        this._AreaerrorPane.show();
    }

    protected clearAreaError() {
        this._AreaerrorPane.text("");
        this._AreaerrorPane.hide();
    }

    private _getCurrentAreaFieldValue(): IPromise<string> {
        var defer = Q.defer();
        WitService.WorkItemFormService.getService().then(
            (service) => {
                service.getFieldValues([this.AreafieldName]).then(
                    (values) => {
                        defer.resolve(values[this.AreafieldName]);
                    },
                    () => {
                        this.showAreaError("Error loading values for field: " + this.AreafieldName);
                    }
                );
            }
        );
        return defer.promise.then();
    }

    //#endregion Area

    //#region Subarea
    public Subareainvalidate(): void {
        if (!this._flushing) {
            this._getCurrentSubareaFieldValue().then(
                (value: string) => {
                    this.setSubareaValue(value);
                }
            );
        }
        this.resize();
    }
    protected getSubareaValue(): string {
        return "";
    }

    protected setSubareaValue(value: string): void {
    }

    protected showSubareaError(error: string): void {
        this._SubareaerrorPane.text(error);
        this._SubareaerrorPane.show();
    }

    protected clearSubareaError() {
        this._SubareaerrorPane.text("");
        this._SubareaerrorPane.hide();
    }

    private _getCurrentSubareaFieldValue(): IPromise<string> {
        var defer = Q.defer();
        WitService.WorkItemFormService.getService().then(
            (service) => {
                service.getFieldValues([this.SubareafieldName]).then(
                    (values) => {
                        defer.resolve(values[this.SubareafieldName]);
                    },
                    () => {
                        this.showSubareaError("Error loading values for field: " + this.SubareafieldName);
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