
import { TCBOMRow, M3BOMRow, IIMRow, BOMError } from '../types';

export const validateBOM = (
  tcData: TCBOMRow[],
  m3Data: M3BOMRow[],
  iimData: IIMRow[]
): { errors: BOMError[], healthScore: number } => {
  const errors: BOMError[] = [];
  const policyMap = new Map<string, string>();

  iimData.forEach(row => {
    policyMap.set(row['Item number'], row['Planning Policy Name']);
  });

  const getPolicy = (item: string) => policyMap.get(item) || 'Purchase';

  // Task 1: Level 1 Validation
  const tcLvl1 = tcData.filter(r => r.Lvl === '1');
  const m3Lvl1 = m3Data.filter(r => r['Rel Lvl'] === '1');
  
  const tcLvl1Map = new Map(tcLvl1.map(r => [r['Part/PA Id Indented'], r]));
  const m3Lvl1Map = new Map(m3Lvl1.map(r => [r['Part no'], r]));

  tcLvl1Map.forEach((row, id) => {
    if (!m3Lvl1Map.has(id)) {
      errors.push({
        id: crypto.randomUUID(),
        task: 'Task 1',
        errorType: 'Missing Level 1',
        parentId: 'Root',
        partId: id,
        description: 'Item exists in Teamcenter Level 1 but missing in M3 ERP structure.',
        priority: 'High',
        actionableFix: 'Use M3 PDS001 to add this item to the parent product structure.'
      });
    } else {
      // Check Quantities at Level 1
      const m3Row = m3Lvl1Map.get(id);
      if (m3Row && row.Qty && m3Row.Qty && parseFloat(row.Qty) !== parseFloat(m3Row.Qty)) {
        errors.push({
          id: crypto.randomUUID(),
          task: 'Quantity Check',
          errorType: 'Quantity Mismatch',
          parentId: 'Root',
          partId: id,
          description: `Quantity mismatch: TC requires ${row.Qty}, M3 lists ${m3Row.Qty}.`,
          priority: 'High',
          actionableFix: 'Update M3 PDS001 quantity to match Engineering intent.'
        });
      }
    }
  });

  // Task 2 & 3: Phantom Analysis
  const m3Structure = new Map<string, Set<string>>();
  m3Data.forEach(r => {
    const parent = r['Prod struct'];
    if (!m3Structure.has(parent)) m3Structure.set(parent, new Set());
    m3Structure.get(parent)?.add(r['Part no']);
  });

  const tcStructure = new Map<string, Set<string>>();
  tcData.forEach(r => {
    const parent = r['Parent Id'];
    if (!tcStructure.has(parent)) tcStructure.set(parent, new Set());
    tcStructure.get(parent)?.add(r['Part/PA Id Indented']);
  });

  const allM3Parents = Array.from(m3Structure.keys());
  allM3Parents.forEach(parent => {
    if (getPolicy(parent) === 'Phantom items') {
      const m3Children = m3Structure.get(parent) || new Set<string>();
      const tcChildren = tcStructure.get(parent) || new Set<string>();

      const missingInM3 = [...tcChildren].filter(x => !m3Children.has(x));
      if (missingInM3.length > 0) {
        errors.push({
          id: crypto.randomUUID(),
          task: 'Task 2',
          errorType: 'Phantom Mismatch',
          parentId: parent,
          partId: missingInM3.join(', '),
          description: 'Phantom parent children sync error.',
          priority: 'Medium',
          actionableFix: 'Synchronize Phantom sub-assembly in M3 PDS001.'
        });
      }

      m3Children.forEach(child => {
        if (getPolicy(child) === 'Phantom items') {
          errors.push({
            id: crypto.randomUUID(),
            task: 'Task 3',
            errorType: 'Policy Violation',
            parentId: parent,
            partId: child,
            description: 'Nesting Phantoms inside Phantoms creates exploding BOM issues in ERP.',
            priority: 'High',
            actionableFix: 'Ensure child part is set to "Purchase" or "Make" in MMS001.'
          });
        }
      });
    }
  });

  const healthScore = Math.max(0, 100 - (errors.length * 2.5));
  return { errors, healthScore };
};
