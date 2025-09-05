let CURRENT_FLOW = "defaultFlow";

class flowManager{
    constructor(){
        this.flow = {};
        this.propertiesService = PropertiesService.getScriptProperties();    
    }

    createFlow(flowName = "defaultFlow"){
        let i = 1;
        while(this.flow[flowName]){
            flowName = `defaultFlow_${i}`;
            i++;
        }
        this.flow[flowName] = {
            steps: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString(),
            stepsCount: 0,
            isSaved: false
        };
        return flowName;
    }

    addStep(flowName = CURRENT_FLOW, methodName, params = []) {
        if (this.flow[flowName]) {
            this.flow[flowName].steps.push({
                sheet: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(),
                method: methodName,
                params: params
            });
            this.flow[flowName].stepsCount++;
            this.flow[flowName].updatedAt = new Date().toISOString();
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    deleteStep(flowName = CURRENT_FLOW) {
        if (this.flow[flowName]) {
            if (this.flow[flowName].steps.length) {
                this.flow[flowName].steps.pop();
                this.flow[flowName].stepsCount--;
                this.flow[flowName].updatedAt = new Date().toISOString();
            } else {
                throw new Error(`No steps to delete in flow ${flowName}.`);
            }
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    executeFlow(flowName = CURRENT_FLOW) {
        if (this.flow[flowName]) {
            const steps = this.flow[flowName].steps;
            for (const step of steps) {
                const tableObj = new table(this.flow[flowName].sheet);
                const {method, params} = step;
                if (typeof tableObj[method] === "function") {
                    tableObj[method](...params);
                } 
                else {
                    throw new Error(`Method ${method} does not exist.`);
                }
            }
            return tableObj;
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    saveFlow(flowName = CURRENT_FLOW) {
        if (this.flow[flowName]) {
            this.propertiesService.setProperty(`${flowName}`, JSON.stringify(this.flow[flowName]));
            this.flow[flowName].isSaved = true;
            return true;
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    loadFlow(flowName) {
        const flowData = this.propertiesService.getProperty(`${flowName}`);
        if (flowData) {
            this.flow[flowName] = JSON.parse(flowData);
            return true;
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    deleteFlow(flowName = CURRENT_FLOW) {
        if (this.flow[flowName]) {
            delete this.flow[flowName];
            this.propertiesService.deleteProperty(`${flowName}`);
            return true;
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    showFlowSteps(flowName = CURRENT_FLOW) {
        if (this.flow[flowName]) {
            let methods = []
            for (const step of this.flow[flowName].steps) {
                methods.push(step.method);
            }
            return methods;
        } else {
            throw new Error(`Flow ${flowName} does not exist.`);
        }
    }

    getAllFlows() {
        const allFlows = this.propertiesService.getProperties();
        const flowNames = Object.keys(allFlows);
        return flowNames;
    }
}

const FLOW_MANAGER = new flowManager();