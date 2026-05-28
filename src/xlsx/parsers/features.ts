import type { RawDocument, ParsedNode } from '../../core/types.js';
import type { AutoFilter, AutoFilterColumn, DataValidation, ConditionalFormat, ConditionalRule, PrintArea } from '../types.js';
import { findChild } from './utils.js';

export function extractAutoFilter(raw: RawDocument, sheetPath: string): AutoFilter | undefined {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return undefined;

  const autoFilterNode = findChild(worksheetXml, 'autoFilter');
  if (!autoFilterNode) return undefined;

  const ref = autoFilterNode.attrs['ref'] || '';
  const columns: AutoFilterColumn[] = [];

  for (const child of autoFilterNode.children) {
    if (typeof child === 'string' || child.tag !== 'filterColumn') continue;
    const colId = parseInt(child.attrs['colId'] || '0', 10);

    const filtersNode = findChild(child, 'filters');
    if (filtersNode) {
      const filters: string[] = [];
      for (const filter of filtersNode.children) {
        if (typeof filter === 'string' || filter.tag !== 'filter') continue;
        if (filter.attrs['val']) filters.push(filter.attrs['val']);
      }
      columns.push({ colId, filters });
    }

    const customFiltersNode = findChild(child, 'customFilters');
    if (customFiltersNode) {
      const customFilter = findChild(customFiltersNode, 'customFilter');
      if (customFilter) {
        columns.push({
          colId,
          customFilter: {
            operator: customFilter.attrs['operator'] || 'equal',
            value: customFilter.attrs['val'] || '',
          },
        });
      }
    }
  }

  return { ref, columns };
}

export function extractDataValidations(raw: RawDocument, sheetPath: string): DataValidation[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const dataValidationsNode = findChild(worksheetXml, 'dataValidations');
  if (!dataValidationsNode) return [];

  const validations: DataValidation[] = [];
  for (const child of dataValidationsNode.children) {
    if (typeof child === 'string' || child.tag !== 'dataValidation') continue;

    const validation: DataValidation = {
      type: child.attrs['type'] || '',
      sqref: child.attrs['sqref'] || '',
    };

    if (child.attrs['operator']) validation.operator = child.attrs['operator'];
    if (child.attrs['allowBlank']) validation.allowBlank = child.attrs['allowBlank'] === '1' || child.attrs['allowBlank'] === 'true';
    if (child.attrs['showErrorMessage']) validation.showErrorMessage = child.attrs['showErrorMessage'] === '1' || child.attrs['showErrorMessage'] === 'true';
    if (child.attrs['errorTitle']) validation.errorTitle = child.attrs['errorTitle'];
    if (child.attrs['error']) validation.error = child.attrs['error'];

    const formula1 = findChild(child, 'formula1');
    if (formula1 && formula1.children.length > 0 && typeof formula1.children[0] === 'string') {
      validation.formula1 = formula1.children[0];
    }

    const formula2 = findChild(child, 'formula2');
    if (formula2 && formula2.children.length > 0 && typeof formula2.children[0] === 'string') {
      validation.formula2 = formula2.children[0];
    }

    validations.push(validation);
  }

  return validations;
}

export function extractConditionalFormats(raw: RawDocument, sheetPath: string): ConditionalFormat[] {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return [];

  const formats: ConditionalFormat[] = [];
  for (const child of worksheetXml.children) {
    if (typeof child === 'string' || child.tag !== 'conditionalFormatting') continue;

    const sqref = child.attrs['sqref'] || '';
    const rules: ConditionalRule[] = [];

    for (const ruleNode of child.children) {
      if (typeof ruleNode === 'string' || ruleNode.tag !== 'cfRule') continue;

      const rule: ConditionalRule = {
        type: ruleNode.attrs['type'] || '',
        priority: parseInt(ruleNode.attrs['priority'] || '0', 10),
      };

      const formula = findChild(ruleNode, 'formula');
      if (formula && formula.children.length > 0 && typeof formula.children[0] === 'string') {
        rule.formula = [formula.children[0]];
      }

      const colorScale = findChild(ruleNode, 'colorScale');
      if (colorScale) {
        const colors: string[] = [];
        for (const colorChild of colorScale.children) {
          if (typeof colorChild === 'string' || colorChild.tag !== 'color') continue;
          if (colorChild.attrs['rgb']) colors.push(colorChild.attrs['rgb']);
        }
        rule.colorScale = { colors };
      }

      rules.push(rule);
    }

    formats.push({ sqref, rules });
  }

  return formats;
}

export function extractPrintArea(raw: RawDocument, sheetPath: string): PrintArea | undefined {
  const worksheetXml = raw.parts.get(sheetPath);
  if (!worksheetXml) return undefined;

  const printOptions = findChild(worksheetXml, 'printOptions');
  const pageMargins = findChild(worksheetXml, 'pageMargins');

  if (!printOptions && !pageMargins) return undefined;

  const printArea: PrintArea = {};

  if (printOptions) {
    if (printOptions.attrs['fitToWidth']) printArea.fitToWidth = parseInt(printOptions.attrs['fitToWidth'], 10);
    if (printOptions.attrs['fitToHeight']) printArea.fitToHeight = parseInt(printOptions.attrs['fitToHeight'], 10);
  }

  if (pageMargins) {
    printArea.pageMargins = {
      top: parseFloat(pageMargins.attrs['top'] || '0'),
      right: parseFloat(pageMargins.attrs['right'] || '0'),
      bottom: parseFloat(pageMargins.attrs['bottom'] || '0'),
      left: parseFloat(pageMargins.attrs['left'] || '0'),
      header: parseFloat(pageMargins.attrs['header'] || '0'),
      footer: parseFloat(pageMargins.attrs['footer'] || '0'),
    };
  }

  return printArea;
}
