import * as WitService from "TFS/WorkItemTracking/Services";
import * as VSSUtilsCore from "VSS/Utils/Core";
import Q = require("q");

export class BaseMultiValueControl {
    /**
     * Field name input for the control
     */
    public fieldName: string;
    public ProductfieldName: string;
    public AreafieldName: string;
    /**
     * The container to hold the control
     */
    protected containerElement: JQuery;
    protected ProductcontainerElement: JQuery;
    protected AreacontainerElement: JQuery;
    /**
     * The container for error message display
     */
    private _errorPane: JQuery;
    private _ProducterrorPane: JQuery;
    private _AreaerrorPane: JQuery;

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

        //#region Family
        this.containerElement = $(".container");
        if (this._showFieldBorder) {
            this.containerElement.addClass("fieldBorder");
        }
       
        this._errorPane = $("<div>").addClass("errorPane").appendTo(this.containerElement);
        var inputs: IDictionaryStringTo<string> = initialConfig.witInputs;

        this.fieldName = inputs["FamilyName"];
        if (!this.fieldName) {
            this.showError("Product input has not been specified");
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

        //#endregion

        //#region Product
        this.ProductcontainerElement = $(".productcontainer");
        if (this._showFieldBorder) {
            this.ProductcontainerElement.addClass("fieldBorder");
        }
       
        this._ProducterrorPane = $("<div>").addClass("errorPane").appendTo(this.ProductcontainerElement);
        var inputs: IDictionaryStringTo<string> = initialConfig.witInputs;

        this.ProductfieldName = inputs["ProductName"];
        if (!this.ProductfieldName) {
            this.showProductError("Product input has not been specified");
        }
        //#endregion

        //#region Area
        this.AreacontainerElement = $(".areacontainer");
        if (this._showFieldBorder) {
            this.AreacontainerElement.addClass("fieldBorder");
        }
       
        this._AreaerrorPane = $("<div>").addClass("errorPane").appendTo(this.AreacontainerElement);
        var inputs: IDictionaryStringTo<string> = initialConfig.witInputs;

        this.AreafieldName = inputs["AreaName"];
        if (!this.AreafieldName) {
            this.showAreaError("Area input has not been specified");
        }
        //#endregion
    }

    /**
     * Initialize a new instance of Control
     */
    public initialize(): void {
        this.invalidate();
    }

    /**
     * Invalidate the control's value
     */
    public invalidate(): void {
        if (!this._flushing) {
            this._getCurrentFieldValue(this.fieldName).then(
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
    protected getValue(): string {
        return "";
    }

    protected setValue(value: string): void {
    }
    protected showError(error: string): void {
        this._errorPane.text(error);
        this._errorPane.show();
    }

    public clearError() {
        this._errorPane.text("");
        this._errorPane.hide();
    }
   
    //#region Product

    protected ProductgetValue(): string {
        return "";
    }
    protected ProductsetValue(value: string): void {
    }
    public Productclear(): void {
    }

    public Productinitialize(): void {
        this.Productinvalidate();
    }
    public Productinvalidate(): void {
        if (!this._flushing) {
            this._getCurrentFieldValue(this.ProductfieldName).then(
                (value: string) => {
                    this.ProductsetValue(value);
                }
            );
        }
        this.resize();
    }
    protected showProductError(error: string): void {
        this._ProducterrorPane.text(error);
        this._ProducterrorPane.show();
    }
    public clearProductError() {
        this._ProducterrorPane.text("");
        this._ProducterrorPane.hide();
    }
    protected Productflush(): void {
        console.log("ProductFlush");
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.ProductfieldName, this.ProductgetValue()).then(
                    (values) => {
                        this._flushing = false;
                    },
                    () => {
                        this._flushing = false;
                        this.showProductError("Error storing the field value");
                    }
                )
            }
        );
    }
    //#endregion

    //#region Area

    protected AreagetValue(): string {
        return "";
    }
    protected AreasetValue(value: string): void {
    }
    public Areaclear(): void {
    }

    public Areainitialize(): void {
        this.Areainvalidate();
    }
    public Areainvalidate(): void {
        if (!this._flushing) {
            this._getCurrentFieldValue(this.AreafieldName).then(
                (value: string) => {
                    this.AreasetValue(value);
                }
            );
        }
        this.resize();
    }
    protected showAreaError(error: string): void {
        this._AreaerrorPane.text(error);
        this._AreaerrorPane.show();
    }
    public clearAreaError() {
        this._AreaerrorPane.text("");
        this._AreaerrorPane.hide();
    }
    protected Areaflush(): void {
        console.log("AreaFlush");
        this._flushing = true;
        WitService.WorkItemFormService.getService().then(
            (service: WitService.IWorkItemFormService) => {
                service.setFieldValue(this.AreafieldName, this.AreagetValue()).then(
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
    //#endregion


    private _getCurrentFieldValue(systemSieldName:string): IPromise<string> {
        var defer = Q.defer();
        WitService.WorkItemFormService.getService().then(
            (service) => {
                service.getFieldValues([systemSieldName]).then(
                    (values) => {
                        defer.resolve(values[systemSieldName]);
                    },
                    () => {
                        this.showError("Error loading values for field: " + systemSieldName)
                    }
                );
            }
        );
        return defer.promise.then();
    }

    protected resize() {
        this._bodyElement = <HTMLBodyElement>document.getElementsByTagName("body").item(0);
        // Cast as any until declarations are updated
        VSS.resize(null, this._bodyElement.offsetHeight);
    }
}