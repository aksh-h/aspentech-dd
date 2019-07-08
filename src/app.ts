import {MultiValueCombo} from "./MultiValueCombo";
import {BaseMultiValueControl} from "./BaseMultiValueControl";
import * as WitExtensionContracts  from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";

// save on ctr + s
$(window).bind("keydown", function (event: JQueryEventObject) {
    if (event.ctrlKey || event.metaKey) {
        if (String.fromCharCode(event.which) === "S") {
            event.preventDefault();
            WorkItemFormService.getService().then((service) => service.beginSaveWorkItem($.noop, $.noop));
        }
    }
});

var control: BaseMultiValueControl;
var stateControl: BaseMultiValueControl;
var CityControl: BaseMultiValueControl;
var provider = () => {
    var ensureControl = () => {
        if (!control) {
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
             control = new MultiValueCombo();
            control.initialize();
        }
        control.invalidate();
    };

    var ensureSateControl  = () => {
        if(!stateControl){
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
            stateControl = new MultiValueCombo();
            stateControl.Stateinitialize();
        }
        stateControl.Stateinvalidate();
    };
    var ensureCityControl  = () => {
        if(!stateControl){
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
            CityControl = new MultiValueCombo();
            CityControl.Cityinitialize();
        }
        CityControl.Cityinvalidate();
    };

    return {
        onLoaded: (args: WitExtensionContracts.IWorkItemLoadedArgs) => {
            ensureControl();
            ensureSateControl();
            ensureCityControl();
        },
        onUnloaded: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
            if (control) {
                control.clear();
                stateControl.clear();
                CityControl.clear();
            }
        },
        onFieldChanged: (args: WitExtensionContracts.IWorkItemFieldChangedArgs) => {
            if (control && args.changedFields[control.fieldName] !== undefined && args.changedFields[control.fieldName] !== null) {
                control.invalidate();
                stateControl.Stateinvalidate();
            } else if (stateControl && args.changedFields[stateControl.statefieldName] !== undefined && args.changedFields[stateControl.statefieldName] !== null) {
                stateControl.Stateinvalidate();
                CityControl.clear();
            }
        }
    };
};

VSS.register(VSS.getContribution().id, provider);

