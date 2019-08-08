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
var Productcontrol: BaseMultiValueControl;
var Areacontrol: BaseMultiValueControl;
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
    var ensureProductcontrol = () =>{
        if(!Productcontrol){
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
            Productcontrol = new MultiValueCombo();
            Productcontrol.Productinitialize();
        }
    }
    var ensureAreacontrol = () =>{
        if(!Areacontrol){
            var inputs: IDictionaryStringTo<string> = VSS.getConfiguration().witInputs;
            var controlType: string = inputs["InputMode"];
            Areacontrol = new MultiValueCombo();
            Areacontrol.Areainitialize();
        }
    }

    return {
        onLoaded: (args: WitExtensionContracts.IWorkItemLoadedArgs) => {
            ensureControl();
            ensureProductcontrol();
        },
        onUnloaded: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
            if (control) {
                control.clear();
            }
            if(Productcontrol){
                Productcontrol.Productclear();
            }
            if(Areacontrol){
                Areacontrol.Areaclear();
            }
        },
        onFieldChanged: (args: WitExtensionContracts.IWorkItemFieldChangedArgs) => {
            if (control && args.changedFields[control.fieldName] !== undefined && args.changedFields[control.fieldName] !== null) {
                control.invalidate();
            }
            else if(Productcontrol && args.changedFields[Productcontrol.ProductfieldName]!==undefined && args.changedFields[Productcontrol.ProductfieldName]!==null){
                Productcontrol.Productinvalidate();
            }
            else if(Areacontrol && args.changedFields[Areacontrol.AreafieldName]!==undefined && args.changedFields[Areacontrol.AreafieldName]!==null){
                Areacontrol.Areainvalidate();
            }
        }
    };
};

VSS.register(VSS.getContribution().id, provider);

