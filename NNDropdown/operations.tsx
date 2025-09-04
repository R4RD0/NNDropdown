import { IInputs } from "./generated/ManifestTypes";
import * as webAPIHelper from "./webAPIHelper";
import { EntityReference, Setting, DropDownOption, DropDownData } from "./interface";
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import * as dropdown from './fluentUIDropdown';

// ----- helpers -----
function getParam(context: ComponentFramework.Context<IInputs>, name: string): string {
  const p = (context.parameters as unknown as Record<string, any>);
  return p?.[name]?.raw ?? "";
}

function log(context: ComponentFramework.Context<IInputs>, msg: string, data?: any) {
  const enabled = (getParam(context, "enableLogging") || "").toLowerCase() === "true";
  if (enabled) console.log(msg, data);
}

export function _labelsForOptions(allOptions: DropDownOption[], selectedKeys: string[]) {
  return allOptions
    .filter(o => selectedKeys.includes(o.key))
    .map(o => o.text)
    .join(", ");
}

// Try to resolve:
// - primaryFieldName (if missing) from target entity metadata
// - relationShipEntityName (if missing) from ManyToMany metadata
async function _resolveMetadata(context: ComponentFramework.Context<IInputs>, setting: Setting) {

  // Resolve primary name attribute if not supplied
  if (!setting.primaryFieldName) {
    const tgtMeta = await context.utils.getEntityMetadata(setting.targetEntityName);
    // PrimaryNameAttribute is standard Dataverse metadata
    setting.primaryFieldName = (tgtMeta as any)?.PrimaryNameAttribute || setting.primaryFieldName;
  }

  // Resolve intersect entity if not supplied
  if (!setting.relationShipEntityName && setting.relationShipName) {
    const primaryMeta = await context.utils.getEntityMetadata(setting.primaryEntityName);
    const targetMeta  = await context.utils.getEntityMetadata(setting.targetEntityName);

    const relNameLower = setting.relationShipName.toLowerCase();

    // Search both sides; match on either SchemaName OR nav property names
    const findM2M = (meta: any) => {
      const rels: any[] = (meta?.ManyToManyRelationships as any[]) || [];
      return rels.find(r =>
        (r.SchemaName && r.SchemaName.toLowerCase() === relNameLower) ||
        (r.Entity1NavigationPropertyName && r.Entity1NavigationPropertyName.toLowerCase() === relNameLower) ||
        (r.Entity2NavigationPropertyName && r.Entity2NavigationPropertyName.toLowerCase() === relNameLower)
      );
    };

    let m2m = findM2M(primaryMeta) || findM2M(targetMeta);

    // Fallback: if a nav property like "..._opportunity" was provided, trim suffix and try again
    if (!m2m) {
      const suffix = "_" + setting.primaryEntityName.toLowerCase();
      if (relNameLower.endsWith(suffix)) {
        const schemaGuess = setting.relationShipName.slice(0, -suffix.length);
        const schemaLower = schemaGuess.toLowerCase();
        const trySchema = (meta: any) => {
          const rels: any[] = (meta?.ManyToManyRelationships as any[]) || [];
          return rels.find(r => (r.SchemaName && r.SchemaName.toLowerCase() === schemaLower));
        };
        m2m = trySchema(primaryMeta) || trySchema(targetMeta);
      }
    }

    if (m2m?.IntersectEntityName) {
      setting.relationShipEntityName = m2m.IntersectEntityName;
    } else {
      log(context, "WARN: Could not auto-resolve IntersectEntityName from metadata. Ensure 'relationshipentityname' is configured.");
    }
  }

  return setting;
}

export function _proccessSetting(context: ComponentFramework.Context<IInputs>) {
  const setting: Setting = {
    // @ts-ignore legacy PCF typings
    primaryEntityId:   (context.page as any)?.entityId ?? "",
    // @ts-ignore legacy PCF typings
    primaryEntityName: (context.page as any)?.entityTypeName ?? "",
    primaryFieldName:         getParam(context, "primaryfieldname"),
    relationShipName:         getParam(context, "relationshipname"),
    relationShipEntityName:   getParam(context, "relationshipentityname"), // optional in manifest
    targetEntityName:         getParam(context, "targetentityname"),
    targetEntityFilter:       getParam(context, "targetentityfilter"),
  };
  return setting;
}

export async function _getAvailableOptions(context: ComponentFramework.Context<IInputs>, setting: Setting) {
  const baseFetchXml = `<fetch><entity name="${setting.targetEntityName}" /></fetch>`;
  const userFetchXml = setting.targetEntityFilter ? setting.targetEntityFilter : "";
  const fetchXml = userFetchXml !== "" ? userFetchXml : baseFetchXml;

  log(context, "FetchXML (all options):", fetchXml);
  const allOptionsSet = await webAPIHelper.retrieveDataFetchXML(context, setting.targetEntityName, fetchXml);

  const allOptions: DropDownOption[] = allOptionsSet.entities.map((entity: any) => ({
    key: entity[`${setting.targetEntityName}id`],
    text: entity[setting.primaryFieldName]
  }));

  allOptions.sort((a, b) => (a.text > b.text ? 1 : (b.text > a.text ? -1 : 0)));
  return allOptions;
}

export async function _currentOptions(context: ComponentFramework.Context<IInputs>, setting: Setting) {
  const fetchXml = `
    <fetch>
      <entity name="${setting.relationShipEntityName}">
        <filter>
          <condition attribute="${setting.primaryEntityName}id" operator="eq" value="${setting.primaryEntityId}" />
        </filter>
      </entity>
    </fetch>`;

  log(context, "FetchXML (selected options):", fetchXml);
  const currentOptionsSet = await webAPIHelper.retrieveDataFetchXML(context, setting.relationShipEntityName, fetchXml);

  return currentOptionsSet.entities.map((entity: any) => entity[`${setting.targetEntityName}id`]);
}

export function _associateRecord(context: ComponentFramework.Context<IInputs>, setting: Setting, targetEntityReference: EntityReference, relatedEntityReference: EntityReference) {
  return new Promise(function (resolve, reject) {
    const req = {
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,
        operationName: "Associate"
      }),
      relationship: setting.relationShipName,
      target: targetEntityReference,
      relatedEntities: [relatedEntityReference]
    };

    log(context, 'Associate request', req);

    // @ts-ignore PCF runtime provides webAPI.execute
    context.webAPI.execute(req).then(resolve, (error: any) => {
      log(context, "Associate error", error);
      reject(error);
    });
  });
}

export function _disAssociateRecord(context: ComponentFramework.Context<IInputs>, setting: Setting, targetEntityReference: EntityReference, relatedEntityReference: EntityReference) {
  return new Promise(function (resolve, reject) {
    const req = {
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,
        operationName: "Disassociate"
      }),
      relationship: setting.relationShipName,
      target: targetEntityReference,
      relatedEntityId: relatedEntityReference.id
    };

    log(context, 'Disassociate request', req);

    // @ts-ignore PCF runtime provides webAPI.execute
    context.webAPI.execute(req).then(resolve, (error: any) => {
      log(context, "Disassociate error", error);
      reject(error);
    });
  });
}

/**
 * Entry point called by index.ts
 */
export async function _execute(
  context: ComponentFramework.Context<IInputs>,
  container: HTMLDivElement,
  updateMirror?: (labels: string) => void
) {
  let setting = _proccessSetting(context);
  log(context, "Settings (raw)", setting);

  // Auto-resolve metadata (primary field + intersect entity)
  setting = await _resolveMetadata(context, setting);
  log(context, "Settings (resolved)", setting);

  if (setting.primaryEntityId) {
    const allOptions = await _getAvailableOptions(context, setting);
    const selectedOptions = await _currentOptions(context, setting);

    const labels = _labelsForOptions(allOptions, selectedOptions);
    if (updateMirror) updateMirror(labels);

    const dropDownData: DropDownData = { allOptions, selectedOptions };

    ReactDOM.render(
      React.createElement(dropdown.NNDropdownControl, {
        context,
        setting,
        dropdowndata: dropDownData,
        onChange: (newSelectedKeys: string[]) => {
          const newLabels = _labelsForOptions(allOptions, newSelectedKeys);
          if (updateMirror) updateMirror(newLabels);
        }
      }),
      container
    );
  } else {
    const msg = <div>This record hasn't been created yet. To enable this control, create the record.</div>;
    ReactDOM.render(msg, container);
  }
}