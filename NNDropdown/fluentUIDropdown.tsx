import * as React from 'react';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { EntityReference, Setting, DropDownData } from "./interface";
import { IInputs } from './generated/ManifestTypes';
import { useState } from 'react';
import * as operations from './operations';

const dropdownStyles: Partial<IDropdownStyles> = {
  title: { border: 'none', backgroundColor: '#F5F5F5' }
};

export const NNDropdownControl: React.FC<{
  context: ComponentFramework.Context<IInputs>,
  setting: Setting,
  dropdowndata: DropDownData,
  onChange?: (selectedKeys: string[]) => void
}> = ({ context, setting, dropdowndata, onChange }) => {

  const targetRef: EntityReference = { entityType: setting.primaryEntityName, id: setting.primaryEntityId };
  const [selectedKeys, setSelectedKeys] = useState<string[]>(dropdowndata.selectedOptions || []);

  const handleChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, index: number): void => {
    if (!item) return;

    let newSelected: string[];

    if (item.selected) {
      newSelected = [...selectedKeys, item.key as string];
      const relatedRef: EntityReference = { entityType: setting.targetEntityName, id: String(item.key) };
      operations._associateRecord(context, setting, targetRef, relatedRef);
    } else {
      newSelected = selectedKeys.filter(key => key !== item.key);
      const relatedRef: EntityReference = { entityType: setting.targetEntityName, id: String(item.key) };
      operations._disAssociateRecord(context, setting, targetRef, relatedRef);
    }

    setSelectedKeys(newSelected);
    if (onChange) onChange(newSelected);
  };

  return (
    <Dropdown
      placeholder="---"
      onChange={handleChange}
      multiSelect
      selectedKeys={selectedKeys}
      options={dropdowndata.allOptions}
      styles={dropdownStyles}
      disabled={context.mode.isControlDisabled}
    />
  );
};