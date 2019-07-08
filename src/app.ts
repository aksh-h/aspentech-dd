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
var AreaControl: BaseMultiValueControl;
var SubareaControl: BaseMultiValueControl;
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
        if(!AreaControl){
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
            AreaControl = new MultiValueCombo();
            AreaControl.Areainitialize();
        }
        AreaControl.Areainvalidate();
    };
    var ensureSubareaControl  = () => {
        if(!AreaControl){
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
            SubareaControl = new MultiValueCombo();
            SubareaControl.Subareainitialize();
        }
        SubareaControl.Subareainvalidate();
    };

    return {
        onLoaded: (args: WitExtensionContracts.IWorkItemLoadedArgs) => {
            ensureControl();
            ensureSateControl();
            ensureSubareaControl();
        },
        onUnloaded: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
            if (control) {
                control.clear();
                AreaControl.clear();
                SubareaControl.clear();
            }
        },
        onFieldChanged: (args: WitExtensionContracts.IWorkItemFieldChangedArgs) => {
            if (control && args.changedFields[control.fieldName] !== undefined && args.changedFields[control.fieldName] !== null) {
                control.invalidate();
                AreaControl.Areainvalidate();
            } else if (AreaControl && args.changedFields[AreaControl.AreafieldName] !== undefined && args.changedFields[AreaControl.AreafieldName] !== null) {
                AreaControl.Areainvalidate();
                SubareaControl.clear();
            }
        }
    };
};

VSS.register(VSS.getContribution().id, provider);

