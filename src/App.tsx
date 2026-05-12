/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Upload, 
  FileText, 
  CheckCircle2, 
  AlertCircle, 
  BarChart3, 
  Search,
  ArrowRight,
  Loader2,
  FileSpreadsheet,
  Download,
  Settings,
  Plus,
  Trash2,
  Database,
  RefreshCw,
  X
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Legend
} from 'recharts';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import * as XLSX from 'xlsx';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Robust numeric parser for Brazilian and US formats
const parseNumeric = (val: any): number => {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (val === undefined || val === null || val === '') return 0;
  let str = String(val).trim();
  str = str.replace(/[R$\s]/g, ''); // Remove currency symbols and spaces

  
  if (str.includes(',') && str.includes('.')) {
    if (str.indexOf('.') < str.indexOf(',')) {
      str = str.replace(/\./g, '').replace(',', '.'); // 1.234,56 -> 1234.56
    } else {
      str = str.replace(/,/g, ''); // 1,234.56 -> 1234.56
    }
  } else if (str.includes(',')) {
    str = str.replace(',', '.'); // 1234,56 -> 1234.56
  }
  
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
};

// String normalization for resilient matching
const normalizeStr = (s: string) => 
  String(s || '').trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/g, "");

// Clean NCM/Natureza codes for uniform comparison
const cleanFiscalCode = (s: any) => String(s || '').trim().replace(/[^0-9]/g, '');


// --- Types ---

interface TaxRule {
  ncm: string;
  natureza: string;
  item: string;
  hasIcms: boolean;
  situacao: 1 | 2 | 3;
  acao: 'Outros Débitos' | 'Estorno' | 'Normal';
}

interface TaxRow {
  id: string;
  ncm: string;
  natureza: string;
  item: string;
  valorContabil: number;
  cstIcms: string;
  baseCalculo: number;
  valorIcms: number;
  outrosDebitos: number;
  estornoDebito: number;
  status: 'Normal' | 'Outros Débitos' | 'Estorno' | 'Pendente';
  matchType?: 'ITEM' | 'NCM' | 'NONE';
}

// --- Initial Rules Database (Local Simulation) ---
const INITIAL_RULES: TaxRule[] = [];

// --- Components ---

const StatCard = ({ title, value, colorClass, icon: Icon }: any) => (
  <div className="glass-card p-6 flex flex-col gap-2">
    <div className="flex justify-between items-start">
      <span className="text-xs font-black text-slate-400 uppercase tracking-widest">{title}</span>
      <div className={cn("p-2 rounded-lg bg-slate-50", colorClass)}>
        <Icon size={18} />
      </div>
    </div>
    <span className="text-2xl font-black text-navy">{value}</span>
  </div>
);

export default function App() {
  const [activeTab, setActiveTab] = useState<'analise' | 'expansao' | 'base'>('analise');
  const [step, setStep] = useState<'upload' | 'processing' | 'results'>('upload');
  const [rules, setRules] = useState<TaxRule[]>(INITIAL_RULES);
  const [processedData, setProcessedData] = useState<TaxRow[]>([]);
  const [pendingItems, setPendingItems] = useState<TaxRule[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [isImportingRules, setIsImportingRules] = useState(false);
  const [ruleImportMode, setRuleImportMode] = useState<'append' | 'replace'>('append');
  const [isDragging, setIsDragging] = useState(false);
  const [companyName, setCompanyName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [rawData, setRawData] = useState<any[]>([]);
  const [triageSelections, setTriageSelections] = useState<Record<string, 1 | 2 | 3>>({});
  const [isSavingTriage, setIsSavingTriage] = useState(false);
  
  const auditFileInputRef = useRef<HTMLInputElement>(null);
  const rulesFileInputRef = useRef<HTMLInputElement>(null);
  const pendingRulesFileInputRef = useRef<HTMLInputElement>(null);
  const dashboardRef = useRef<HTMLDivElement>(null);

  const fetchRules = async () => {
    const saved = localStorage.getItem('tax_rules_cp');
    if (saved) {
      try {
        setRules(JSON.parse(saved));
      } catch (err) {
        console.error("Erro ao carregar regras locais:", err);
      }
    }
  };

  useEffect(() => {
    fetchRules();
  }, []);

  // Persist rules to localStorage whenever they change
  useEffect(() => {
    if (rules.length >= 0) {
      localStorage.setItem('tax_rules_cp', JSON.stringify(rules));
    }
  }, [rules]);


  // Process data based on current rules
  const processTaxData = (rawData: any[], currentRules: TaxRule[]) => {
    if (!Array.isArray(rawData)) return [];
    
    // Pre-index rules for O(1) lookup
    const itemRules = new Map<string, TaxRule>();

    currentRules.forEach(r => {
      const itemKey = normalizeStr(r.item) + '_' + r.hasIcms;
      // Em caso de duplicatas pontuais, a última prevalece
      itemRules.set(itemKey, r);
    });

    const findValueInRow = (r: any, possibleNames: string[]) => {
      const keys = Object.keys(r);
      const cleanNames = possibleNames.map(n => n.toUpperCase().replace(/\s/g, "").normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
      const key = keys.find(k => {
        const cleanK = String(k).toUpperCase().replace(/\s/g, "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        return cleanNames.includes(cleanK);
      });
      return key ? r[key] : undefined;
    };

    return rawData.map((row, index) => {
      if (!row || Object.keys(row).length === 0) return null;
      
      const itemRaw = findValueInRow(row, ['ITEM', 'PRODUTO', 'DESCRICAO']);
      const itemName = String(itemRaw || '').trim().toUpperCase();
      
      // Ignore empty items or rows that look like totals/summaries
      if (!itemName || itemName === 'TOTAL' || itemName === 'SUBTOTAL' || itemName.includes('TOTAL DA NOTA')) {
        return null;
      }

      const valorContabilRaw = findValueInRow(row, ['VALOR CONTABIL', 'VALOR_CONTABIL', 'CONTABIL', 'VALOR TOTAL']);
      const valorContabil = parseNumeric(valorContabilRaw);
      
      // Robust ICMS detection in audit file (numeric or text)
      const valorIcmsRaw = findValueInRow(row, ['VALOR ICMS', 'ICMS', 'VALOR_ICMS', 'VALOR ICM']);
      const temIcmsRaw = findValueInRow(row, ['TEM ICMS', 'ICMS', 'COM ICMS', 'TEM_ICMS']);
      
      let rowHasIcms = parseNumeric(valorIcmsRaw) > 0;
      if (!rowHasIcms && temIcmsRaw) {
        const cleanTem = String(temIcmsRaw).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        rowHasIcms = ['SIM', 'S', 'TRUE'].includes(cleanTem);
      }

      const valorIcms = parseNumeric(valorIcmsRaw);
      const normalizedItem = normalizeStr(itemName);
      const rowNcmRaw = findValueInRow(row, ['NCM', 'CODIGO NCM', 'CÓDIGO NCM']);
      const rowNcm = String(rowNcmRaw || '').trim();
      const rowNaturezaRaw = findValueInRow(row, ['NATUREZA', 'CFOP', 'COD PRODUTO', 'REF']);
      const rowNatureza = String(rowNaturezaRaw || '').trim();

      // Find matching rule by Normalized Item Name AND ICMS presence
      const itemKey = normalizedItem + '_' + rowHasIcms;
      let matchingRule = itemRules.get(itemKey);
      let matchType: 'ITEM' | 'NCM' | 'NONE' = matchingRule ? 'ITEM' : 'NONE';

      let status: 'Normal' | 'Outros Débitos' | 'Estorno' | 'Pendente' = 'Normal';
      let outrosDebitos = 0;
      let estornoDebito = 0;

      if (matchingRule) {
        if (matchingRule.situacao === 2) {
          // Priority: If the item already has ICMS highlighted, treat as Normal even if rule says Outros Débitos
          if (valorIcms > 0) {
            status = 'Normal';
          } else {
            status = 'Outros Débitos';
            outrosDebitos = valorContabil * 0.205; 
          }
        } else if (matchingRule.situacao === 3) {
          status = 'Estorno';
          estornoDebito = valorIcms;
        } else {
          status = 'Normal';
        }
      } else {
        status = 'Pendente';
      }


      const cstRaw = findValueInRow(row, ['CST', 'CST ICMS', 'CST_ICMS']);
      const bcRaw = findValueInRow(row, ['BASE CALCULO', 'BASE CALCULO ICMS', 'BASE CALCULO ICM', 'BASE_CALCULO']);

      return {
        id: `row-${index}`,
        ncm: rowNcm,
        natureza: String(findValueInRow(row, ['NATUREZA', 'CFOP', 'NAT']) || ''),
        item: itemName,
        valorContabil,
        cstIcms: cstRaw !== undefined && cstRaw !== null ? String(cstRaw) : '',
        baseCalculo: bcRaw !== undefined && bcRaw !== null ? Number(bcRaw) : 0,
        valorIcms: status === 'Estorno' ? 0 : (status === 'Outros Débitos' ? outrosDebitos : valorIcms),
        outrosDebitos,
        estornoDebito,
        status,
        matchType
      };
    }).filter(Boolean) as TaxRow[];
  };

  const clearRules = async () => {
    if (confirm("Tem certeza que deseja apagar todas as regras do banco local? Esta ação não pode ser desfeita.")) {
      setRules([]);
    }
  };

  const exportRules = () => {
    if (rules.length === 0) {
      alert("O banco de regras está vazio.");
      return;
    }
    const exportData = rules.map(r => ({
      'NCM': r.ncm,
      'NATUREZA': r.natureza,
      'ITEM': r.item,
      'VALOR ICMS': r.hasIcms ? 'SIM' : 'NÃO',
      'SITUAÇÃO': r.situacao
    }));
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Regras");
    XLSX.writeFile(workbook, "Banco_de_Regras_ICMS_CP.xlsx");
  };

  const handleFile = (file: File, type: 'audit' | 'rules', mode: 'append' | 'replace' = 'append') => {
    setError(null);
    if (type === 'audit') setStep('processing');
    else setIsImportingRules(true);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);
        
        if (type === 'audit') {
          setCompanyName(file.name.replace(/\.[^/.]+$/, ""));
          setTimeout(() => {
            try {
              setRawData(data);
              const processed = processTaxData(data, rules);
              setProcessedData(processed);
              
              // Identify unique pending items for the "Export Pending" feature
              const pendingRows = processed.filter(r => r.status === 'Pendente');
              const uniquePending = new Map<string, TaxRule>();
              pendingRows.forEach(r => {
                const key = r.item;
                if (!uniquePending.has(key)) {
                  uniquePending.set(key, {
                    ncm: r.ncm,
                    natureza: r.natureza,
                    item: r.item,
                    hasIcms: r.valorIcms > 0,
                    situacao: 1, // Default for analyst to change
                    acao: 'Normal'
                  });
                }
              });
              setPendingItems(Array.from(uniquePending.values()));
              
              setStep('results');
            } catch (err) {
              console.error(err);
              setError("Erro ao processar os dados da planilha. Verifique as colunas.");
              setStep('upload');
            }
          }, 1000);
        } else {
          setTimeout(() => {
            try {
              const rawRules = data as any[];
              const rulesMap = new Map<string, TaxRule>();
              
              let foundSituacaoColumn = false;

              rawRules.forEach(row => {
                const keys = Object.keys(row);
                const findValue = (possibleNames: string[], partialNames: string[] = []) => {
                  const cleanNames = possibleNames.map(n => String(n).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s/g, ""));
                  const cleanPartials = partialNames.map(n => String(n).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s/g, ""));
                  
                  const key = keys.find(k => {
                    const cleanK = String(k).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s/g, "");
                    if (cleanNames.includes(cleanK)) return true;
                    if (cleanPartials.some(p => cleanK.includes(p))) return true;
                    return false;
                  });
                  return key ? row[key] : undefined;
                };

                const itemRaw = findValue(['ITEM', 'NOME', 'PRODUTO', 'DESCRICAO', 'DESCRIÇÃO']);
                const item = String(itemRaw || '').trim().toUpperCase();
                if (!item) return;

                const ncmRaw = findValue(['NCM', 'CÓDIGO NCM', 'CÓD. NCM', 'NATUREZA', 'NATURE', 'NAT']);
                let ncm = String(ncmRaw || '').trim();
                // If it's stored as scientific notation or something, stringify cleanly
                if (typeof ncmRaw === 'number') {
                  ncm = ncmRaw.toString();
                }

                const naturezaRaw = findValue(['NATUREZA', 'CFOP', 'NATURE']);
                const natureza = String(naturezaRaw || '').trim().toUpperCase();
                const situacaoRaw = findValue(
                  ['SITUACAO', 'SITUÇÃO', 'SITUCAO', 'STATUS', 'SIT', 'ACAO', 'AÇÃO', 'CLASSIFICACAO', 'TIPO', 'CST'], 
                  ['SITU', 'CLASS', 'TRIBUT']
                );
                
                let situacao: 1 | 2 | 3 = 1;
                if (situacaoRaw !== undefined && situacaoRaw !== null) {
                  foundSituacaoColumn = true;
                  let valStr = String(situacaoRaw).trim().toUpperCase();
                  valStr = valStr.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                  
                  if (valStr === '3' || valStr === '3.0' || valStr === '3,0') situacao = 3;
                  else if (valStr === '2' || valStr === '2.0' || valStr === '2,0') situacao = 2;
                  else if (valStr.includes('3') || valStr.includes('ESTORNO')) situacao = 3;
                  else if (valStr.includes('2') || valStr.includes('OUTRO') || valStr.includes('20,5') || valStr.includes('20.5')) situacao = 2;
                  else situacao = 1;
                }

                const valorIcmsRaw = findValue(
                  ['VALOR ICMS', 'ICMS', 'VALORICMS', 'TEM ICMS', 'CONDICAO ICMS', 'STATUS ICMS', 'DESTACA ICMS', 'BASE CALCULO ICM', 'BASE CALCULO ICMS', 'BASE CÁLCULO ICM']
                );
                let hasIcms = false;
                if (typeof valorIcmsRaw === 'string') {
                  const cleanVal = valorIcmsRaw.trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                  hasIcms = cleanVal === 'SIM' || cleanVal === 'S' || cleanVal === 'TRUE' || cleanVal.includes('SIM') || cleanVal.includes('COM ICMS') || parseNumeric(cleanVal) > 0;
                } else if (valorIcmsRaw !== undefined && valorIcmsRaw !== null) {
                  hasIcms = parseNumeric(valorIcmsRaw) > 0;
                }


                // Determine action from Situacao
                let acao: 'Outros Débitos' | 'Estorno' | 'Normal' = 'Normal';
                if (situacao === 2) acao = 'Outros Débitos';
                else if (situacao === 3) acao = 'Estorno';

                const key = normalizeStr(item) + '_' + hasIcms;
                // Sobrescreve com a regra mais recente (a última lida na planilha)
                rulesMap.set(key, { ncm, natureza, item, hasIcms, situacao, acao });
              });

              // Filter out rules that don't have situations > 1 if they are "duplicates" 
              // Actually, no, we just keep the highest situation for each unique item name.
              const newRules = Array.from(rulesMap.values());

              fetch('/api/rules', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ rules: newRules, mode })
              })
              .then(res => {
                if (!res.ok) throw new Error('Failed to save rules');
                return res.json();
              })
              // Update rules in state
              if (mode === 'replace') {
                setRules(newRules);
                setIsImportingRules(false);
                let msg = `${newRules.length} regras substituídas com sucesso!`;
                if (!foundSituacaoColumn) {
                  msg += "\n\n⚠️ AVISO: A coluna de 'SITUAÇÃO' não foi encontrada. Todos os itens foram salvos como Situação 1.";
                }
                alert(msg);
              } else {
                // APPEND com regra de unicidade: Chave = Nome + ICMS
                // - Se Nome+ICMS NÃO existe → ADICIONA
                // - Se Nome+ICMS existe com Situação DIFERENTE → ATUALIZA (nova Situação prevalece)
                // - Se Nome+ICMS existe com Situação IGUAL → IGNORA (sem alteração)
                let addedCount = 0;
                let updatedCount = 0;
                let skippedCount = 0;
                setRules(prev => {
                  const merged = new Map<string, TaxRule>();
                  prev.forEach(r => merged.set(normalizeStr(r.item) + '_' + r.hasIcms, r));
                  newRules.forEach(r => {
                    const key = normalizeStr(r.item) + '_' + r.hasIcms;
                    const existing = merged.get(key);
                    if (!existing) {
                      merged.set(key, r);
                      addedCount++;
                    } else if (existing.situacao !== r.situacao) {
                      merged.set(key, r);
                      updatedCount++;
                    } else {
                      skippedCount++;
                    }
                  });
                  return Array.from(merged.values());
                });
                setIsImportingRules(false);
                let msg = `Importação concluída!\n✅ ${addedCount} regras novas adicionadas.\n🔄 ${updatedCount} regras atualizadas (mesma chave, Situação diferente).\n⏭️ ${skippedCount} regras idênticas ignoradas.`;
                if (!foundSituacaoColumn) {
                  msg += "\n\n⚠️ AVISO: A coluna de 'SITUAÇÃO' não foi encontrada. Novos itens salvos como Situação 1.";
                }
                alert(msg);
              }

            } catch (err) {
              console.error(err);
              setError("Erro ao interpretar arquivo de regras. Verifique o formato.");
              setIsImportingRules(false);
            }
          }, 1000);
        }
      } catch (err) {
        console.error(err);
        setError("Erro ao ler o arquivo. Certifique-se de que é um Excel válido.");
        if (type === 'audit') setStep('upload');
        setIsImportingRules(false);
      }
    };
    reader.onerror = () => {
      setError("Erro na leitura do arquivo.");
      if (type === 'audit') setStep('upload');
      setIsImportingRules(false);
    };
    reader.readAsBinaryString(file);
  };

  const handleAuditFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file, 'audit');
    e.target.value = '';
  };

  const handleRulesFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file, 'rules', ruleImportMode);
    // Reset input
    e.target.value = '';
  };

  const handleBulkPendingImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setError(null);
    setIsSavingTriage(true);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        const newRules: TaxRule[] = [];
        data.forEach(row => {
          const findValue = (possibleNames: string[]) => {
            const keys = Object.keys(row);
            const cleanNames = possibleNames.map(n => 
              String(n).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s/g, "")
            );
            const key = keys.find(k => {
              const cleanK = String(k).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s/g, "");
              return cleanNames.includes(cleanK);
            });
            return key ? row[key] : undefined;
          };

          const item = String(findValue(['ITEM', 'PRODUTO']) || '').trim().toUpperCase();
          const situacaoRaw = findValue(['SITUACAO', 'SITUÇÃO', 'SIT']);
          const hasIcmsRaw = findValue(['TEM ICMS', 'ICMS', 'COM ICMS']);

          if (!item || situacaoRaw === undefined) return;

          let situacao: 1 | 2 | 3 = 1;
          const s = String(situacaoRaw).trim();
          if (s === '3') situacao = 3;
          else if (s === '2') situacao = 2;

          let hasIcms = false;
          if (typeof hasIcmsRaw === 'string') {
            const clean = hasIcmsRaw.trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            hasIcms = ['SIM', 'S', 'TRUE'].includes(clean);
          } else if (hasIcmsRaw !== undefined) {
            hasIcms = parseNumeric(hasIcmsRaw) > 0;
          }

          // Try to find NCM and Natureza from existing data if possible, else empty
          const existingPending = pendingItems.find(p => normalizeStr(p.item) === normalizeStr(item) && p.hasIcms === hasIcms);

          newRules.push({
            ncm: existingPending?.ncm || String(findValue(['NCM', 'CODIGO NCM']) || ''),
            natureza: existingPending?.natureza || String(findValue(['NATUREZA', 'CFOP', 'NAT']) || ''),
            item,
            hasIcms,
            situacao,
            acao: situacao === 2 ? 'Outros Débitos' : situacao === 3 ? 'Estorno' : 'Normal'
          });

        });

        if (newRules.length === 0) {
          alert("Nenhuma regra válida encontrada na planilha.");
          setIsSavingTriage(false);
          return;
        }

        // Update local rules
          setRules(prev => {
            const merged = new Map<string, TaxRule>();
            prev.forEach(r => merged.set(normalizeStr(r.item) + '_' + r.hasIcms, r));
            newRules.forEach(r => merged.set(normalizeStr(r.item) + '_' + r.hasIcms, r));
            return Array.from(merged.values());
          });
        
        // Final re-process of active analysis if current data exists
        if (rawData.length > 0) {
          const currentAllRules = Array.from(new Map([
            ...rules.map(r => [normalizeStr(r.item) + '_' + r.hasIcms, r]),
            ...newRules.map(r => [normalizeStr(r.item) + '_' + r.hasIcms, r])
          ].reverse()).values()) as TaxRule[];
          
          const newlyProcessed = processTaxData(rawData, currentAllRules);
          setProcessedData(newlyProcessed);
          
          const pendingRows = newlyProcessed.filter(r => r.status === 'Pendente');
          const uniquePending = new Map<string, TaxRule>();
          pendingRows.forEach(r => {
            const key = normalizeStr(r.item);
            if (!uniquePending.has(key)) {
              uniquePending.set(key, {
                ncm: r.ncm,
                natureza: r.natureza,
                item: r.item,
                hasIcms: r.valorIcms > 0,
                situacao: 1,
                acao: 'Normal'
              });
            }
          });
          setPendingItems(Array.from(uniquePending.values()));
        }
        
        alert(`${newRules.length} regras importadas com sucesso!`);
      } catch (err) {
        console.error(err);
        setError("Erro ao processar planilha de pendentes.");
      } finally {
        setIsSavingTriage(false);
        e.target.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = () => {
    setIsDragging(false);
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file, 'audit');
  };

  const totals = useMemo(() => {
    return processedData.reduce((acc, curr) => {
      acc.contabil += curr.valorContabil;
      acc.icms += curr.valorIcms;
      acc.outros += curr.outrosDebitos;
      acc.estorno += curr.estornoDebito;
      
      const status = curr.status;
      if (!acc.byStatus[status]) {
        acc.byStatus[status] = { count: 0, contabil: 0, icms: 0, outros: 0, estorno: 0 };
      }
      acc.byStatus[status].count += 1;
      acc.byStatus[status].contabil += curr.valorContabil;
      acc.byStatus[status].icms += curr.valorIcms;
      acc.byStatus[status].outros += curr.outrosDebitos;
      acc.byStatus[status].estorno += curr.estornoDebito;
      
      return acc;
    }, { 
      contabil: 0, icms: 0, outros: 0, estorno: 0, 
      byStatus: {} as Record<string, { count: number, contabil: number, icms: number, outros: number, estorno: number }> 
    });
  }, [processedData]);

  const filteredData = useMemo(() => {
    return processedData
      .filter(row => 
        row.item.toLowerCase().includes(searchTerm.toLowerCase()) || 
        row.ncm.toLowerCase().includes(searchTerm.toLowerCase())
      )
      .slice(0, 100);
  }, [processedData, searchTerm]);

  const formatCurrency = (val: number) => {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
  };

  const removeRule = (item: string) => {
    setRules(prev => prev.filter(r => normalizeStr(r.item) !== normalizeStr(item)));
  };

  const exportToExcel = () => {
    if (processedData.length === 0) return;
    
    const exportData = processedData.map(row => ({
      'NCM': row.ncm,
      'Natureza': row.natureza,
      'Item': row.item,
      'Valor Contábil': row.valorContabil,
      'CST ICMS': row.cstIcms,
      'Base Cálculo ICM': row.baseCalculo,
      'Valor ICMS': row.valorIcms,
      'Outros Débitos': row.outrosDebitos,
      'Estorno Débito': row.estornoDebito,
      'STATUS': row.status
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Apuração_Completa");
    
    const fileName = `${companyName || 'Apuração'}_TRIBUTACAO_COMPLETA.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  const exportPendingItems = () => {
    if (pendingItems.length === 0) {
      alert("Não há itens pendentes para exportar.");
      return;
    }

    const exportData = pendingItems.map(r => ({
      'NCM': r.ncm,
      'NATUREZA': r.natureza,
      'ITEM': r.item,
      'TEM ICMS': r.hasIcms ? 'SIM' : 'NÃO',
      'SITUAÇÃO': '' // Analista preencherá (1, 2 ou 3)
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    
    // Auto-adjust column widths
    const wscols = [
      {wch: 15}, // NCM
      {wch: 15}, // NATUREZA
      {wch: 50}, // ITEM
      {wch: 12}, // TEM ICMS
      {wch: 12}  // SITUAÇÃO
    ];
    worksheet['!cols'] = wscols;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Itens_Pendentes");
    
    const fileName = `${companyName || 'Apuração'}_ITENS_PENDENTES.xlsx`;
    XLSX.writeFile(workbook, fileName);
    alert(`Exportado ${pendingItems.length} itens pendentes. Preencha a coluna SITUAÇÃO (1, 2 ou 3) e re-importe no Banco de Regras.`);
  };


  const exportToPDF = () => {
    const originalTitle = document.title;
    document.title = `${companyName || 'Apuração'}_TRIBUTACAO_COMPLETA`;
    window.print();
    setTimeout(() => {
      document.title = originalTitle;
    }, 1000);
  };

  const saveTriageAndRefresh = async () => {
    const itemsToSave = Object.keys(triageSelections);
    if (itemsToSave.length === 0) {
      alert("Nenhuma alteração para salvar.");
      return;
    }

    setIsSavingTriage(true);
    try {
      const newRulesBatch = itemsToSave.map(key => {
        const [itemName, hasIcmsStr] = key.split('|');
        const pending = pendingItems.find(p => p.item === itemName);
        return {
          ncm: pending?.ncm || "",
          natureza: pending?.natureza || "",
          item: itemName,
          hasIcms: hasIcmsStr === 'true',
          situacao: triageSelections[key],
          acao: triageSelections[key] === 2 ? 'Outros Débitos' : triageSelections[key] === 3 ? 'Estorno' : 'Normal'
        };
      });

      // Update rules in state
      const nextRules = [...rules];
      newRulesBatch.forEach(nr => {
        const idx = nextRules.findIndex(r => 
          (normalizeStr(r.item) + '_' + r.hasIcms) === (normalizeStr(nr.item) + '_' + nr.hasIcms)
        );
        if (idx >= 0) nextRules[idx] = nr;
        else nextRules.push(nr);
      });
      setRules(nextRules);
      
      const newlyProcessed = processTaxData(rawData, nextRules);
      setProcessedData(newlyProcessed);
      
      // Re-calculate pending items
      const pendingRows = newlyProcessed.filter(r => r.status === 'Pendente');
      const uniquePending = new Map<string, TaxRule>();
      pendingRows.forEach(r => {
        const key = normalizeStr(r.item);
        if (!uniquePending.has(key)) {
          uniquePending.set(key, {
            ncm: r.ncm,
            natureza: r.natureza,
            item: r.item,
            hasIcms: r.valorIcms > 0,
            situacao: 1,
            acao: 'Normal'
          });
        }
      });
      setPendingItems(Array.from(uniquePending.values()));
      setTriageSelections({});
      
      alert("Análise recalculada com novas regras!");
    } catch (err) {
      console.error(err);
      alert("Erro ao salvar regras e atualizar análise.");
    } finally {
      setIsSavingTriage(false);
    }
  };
  return (
    <div className="min-h-screen flex flex-col bg-slate-50">
      {/* Hidden File Inputs */}
      <input 
        type="file" 
        ref={auditFileInputRef} 
        onChange={handleAuditFileUpload} 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
      />
      <input 
        type="file" 
        ref={rulesFileInputRef} 
        onChange={handleRulesFileUpload} 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
      />
      <input 
        type="file" 
        ref={pendingRulesFileInputRef} 
        onChange={handleBulkPendingImport} 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
      />

      {/* Header */}
      <header className="bg-navy text-white px-8 py-5 flex justify-between items-center sticky top-0 z-50 shadow-xl">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-white/10 rounded-2xl flex items-center justify-center border border-white/20 shadow-inner">
            <Database size={28} className="text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-black tracking-tighter">ANÁLISE ICMS <span className="text-blue-400">CP</span></h1>
            <p className="text-[10px] uppercase tracking-[0.3em] opacity-50 font-bold">Auditoria de Situações ICMS</p>
          </div>
        </div>
        <nav className="flex items-center gap-8">
          <button 
            onClick={() => {
              setActiveTab('analise');
              setStep('upload');
              setError(null);
            }}
            className={cn(
              "text-sm font-black tracking-widest uppercase transition-all hover:text-blue-400",
              activeTab === 'analise' ? "text-blue-400" : "text-white/70"
            )}
          >
            Análise ICMS
          </button>
          <button 
            onClick={() => setActiveTab('expansao')}
            className={cn(
              "text-sm font-black tracking-widest uppercase transition-all hover:text-blue-400",
              activeTab === 'expansao' ? "text-blue-400" : "text-white/70"
            )}
          >
            Expansão do Banco de Regras
          </button>
          <button 
            onClick={() => setActiveTab('base')}
            className={cn(
              "text-sm font-black tracking-widest uppercase transition-all hover:text-blue-400 flex items-center gap-2",
              activeTab === 'base' ? "text-blue-400" : "text-white/70"
            )}
          >
            <Settings size={16} /> Banco de Regras
          </button>
        </nav>
      </header>

      <main className="flex-1 max-w-[1600px] mx-auto w-full p-8">
        <AnimatePresence mode="wait">
          {activeTab === 'analise' && (
            <motion.div 
              key="analise"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
            >
              {step === 'upload' && (
                <motion.div 
                  key="upload"
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 1.05 }}
                  className="flex flex-col items-center justify-center py-24"
                >
                  <div className="text-center mb-16">
                    <h2 className="text-5xl font-black text-navy mb-6 tracking-tight">Análise de ICMS</h2>
                    <p className="text-slate-400 text-xl max-w-2xl mx-auto font-medium">
                      Importe sua planilha de apuração. O sistema confrontará os itens com o banco de regras e validará o destaque do <span className="text-navy font-bold">ICMS</span>.
                    </p>
                  </div>

                  <div 
                    onClick={() => auditFileInputRef.current?.click()}
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onDrop={onDrop}
                    className="w-full max-w-2xl group cursor-pointer"
                  >
                    <div className={cn(
                      "relative p-12 border-4 border-dashed rounded-[40px] bg-white transition-all duration-500 flex flex-col items-center",
                      isDragging ? "border-blue-400 bg-blue-50 shadow-2xl scale-105" : "border-slate-200 hover:border-navy hover:shadow-2xl"
                    )}>
                      <div className={cn(
                        "w-24 h-24 rounded-3xl flex items-center justify-center mb-8 transition-all duration-500",
                        isDragging ? "bg-blue-400 rotate-12" : "bg-slate-50 group-hover:bg-navy group-hover:rotate-6"
                      )}>
                        <Upload className={cn(
                          "transition-colors",
                          isDragging ? "text-white" : "text-slate-300 group-hover:text-white"
                        )} size={48} />
                      </div>
                      <h3 className="text-2xl font-black text-navy mb-2">
                        {isDragging ? "Solte para analisar" : "Anexar Planilha de Movimentação"}
                      </h3>
                      <p className="text-slate-400 font-bold tracking-wide">Arraste e solte ou clique para selecionar</p>
                      <p className="text-slate-300 text-xs mt-2 font-bold uppercase tracking-widest">XLSX / CSV / XLS</p>
                      
                      <div className="absolute -bottom-6 left-1/2 -translate-x-1/2 bg-navy text-white px-8 py-3 rounded-full font-black text-sm shadow-xl group-hover:scale-110 transition-transform">
                        INICIAR ANÁLISE
                      </div>
                    </div>
                  </div>

                  {/* Manual Order Guide */}
                  <div className="mt-16 p-8 bg-amber-50 border border-amber-100 rounded-3xl max-w-2xl w-full text-center">
                    <h4 className="text-amber-800 font-black text-xs uppercase tracking-widest mb-4 flex items-center justify-center gap-2">
                      <FileSpreadsheet size={16} /> Ordem Obrigatória das Colunas
                    </h4>
                    <div className="flex flex-wrap justify-center gap-2">
                      {[
                        { col: 'A', name: 'NCM' },
                        { col: 'B', name: 'Natureza' },
                        { col: 'C', name: 'Item' },
                        { col: 'D', name: 'Valor Contábil' },
                        { col: 'E', name: 'CST ICMS' },
                        { col: 'F', name: 'Base Cálculo' },
                        { col: 'G', name: 'Valor ICMS' }
                      ].map(col => (
                        <div key={col.col} className="flex flex-col gap-1">
                          <span className="text-[10px] font-black text-amber-600/50">{col.col}</span>
                          <span className="px-3 py-2 bg-white border border-amber-200 rounded-xl text-xs font-black text-amber-700 shadow-sm whitespace-nowrap">
                            {col.name}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {error && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="mt-12 p-4 bg-red-50 border border-red-100 rounded-2xl flex items-center gap-3 text-red-600 font-bold text-sm"
                    >
                      <AlertCircle size={20} />
                      {error}
                      <button onClick={() => setError(null)} className="ml-auto hover:bg-red-100 p-1 rounded-lg">
                        <X size={16} />
                      </button>
                    </motion.div>
                  )}
                </motion.div>
              )}

              {step === 'processing' && (
                <motion.div 
                  key="processing"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="flex flex-col items-center justify-center py-40"
                >
                  <div className="relative">
                    <Loader2 className="w-32 h-32 text-navy animate-spin opacity-20" strokeWidth={1} />
                    <div className="absolute inset-0 flex items-center justify-center">
                      <motion.div
                        animate={{ scale: [1, 1.2, 1], rotate: [0, 10, -10, 0] }}
                        transition={{ repeat: Infinity, duration: 2 }}
                      >
                        <FileText size={48} className="text-navy" />
                      </motion.div>
                    </div>
                  </div>
                  <h3 className="text-3xl font-black text-navy mt-12 mb-4">Validando Situações Tributárias</h3>
                  <p className="text-slate-400 font-bold tracking-[0.2em] uppercase text-xs">Mapeando itens idênticos no banco de dados...</p>
                </motion.div>
              )}

              {step === 'results' && (
                <motion.div 
                  key="results"
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="space-y-8"
                >
                  <div className="flex justify-between items-center no-print">
                    <button 
                      onClick={() => setStep('upload')}
                      className="text-navy font-black text-xs uppercase tracking-widest flex items-center gap-2 hover:translate-x-[-4px] transition-transform"
                    >
                      <ArrowRight className="rotate-180" size={16} /> Novo Upload
                    </button>
                    <div className="flex gap-4">
                      <button 
                        onClick={exportToPDF}
                        className="px-6 py-3 bg-white border border-slate-200 rounded-xl font-black text-xs text-navy uppercase tracking-widest hover:bg-slate-50 transition-all flex items-center gap-2"
                      >
                        <Download size={16} /> Exportar PDF
                      </button>
                      {pendingItems.length > 0 && (
                        <button 
                          onClick={exportPendingItems}
                          className="px-6 py-3 bg-amber-500 text-white rounded-xl font-black text-xs uppercase tracking-widest hover:bg-amber-600 shadow-lg transition-all flex items-center gap-2"
                        >
                          <FileSpreadsheet size={16} /> Exportar Pendentes ({pendingItems.length})
                        </button>
                      )}
                      <button 
                        onClick={exportToExcel}
                        className="px-6 py-3 bg-navy text-white rounded-xl font-black text-xs uppercase tracking-widest hover:bg-navy-light shadow-lg transition-all flex items-center gap-2"
                      >
                        <FileSpreadsheet size={16} /> Exportar Excel
                      </button>
                    </div>
                  </div>

                  {/* Capturable Dashboard Area */}
                  <div ref={dashboardRef} className="p-8 bg-slate-50 rounded-[40px]">
                    {companyName && (
                      <div className="text-center pb-8">
                        <h2 className="text-4xl font-black text-navy uppercase tracking-tighter">{companyName}</h2>
                        <p className="text-slate-400 font-bold uppercase tracking-[0.3em] text-[10px] mt-1">Dashboad de Apuração Mensal - ICMS</p>
                      </div>
                    )}

                    <div className="space-y-6">
                      {/* Summary Cards */}
                      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                        <StatCard title="Valor Contábil Total" value={formatCurrency(totals.contabil)} icon={FileText} colorClass="text-slate-400" />
                        <StatCard title="ICMS Apurado" value={formatCurrency(totals.icms)} icon={BarChart3} colorClass="text-blue-500" />
                        <StatCard title="Outros Débitos" value={formatCurrency(totals.outros)} icon={Plus} colorClass="text-amber-500" />
                        <StatCard title="Estorno de Débito" value={formatCurrency(totals.estorno)} icon={Trash2} colorClass="text-red-500" />
                      </div>

                      {/* Dashboard by Status */}
                      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                        {['Normal', 'Outros Débitos', 'Estorno', 'Pendente'].map((status) => {
                          const statusData = totals.byStatus[status] || { count: 0, contabil: 0, icms: 0, outros: 0, estorno: 0 };
                          const statusLabels = {
                            'Normal': { name: "Status: Normal", color: "bg-emerald-100 text-emerald-700", border: "border-emerald-200" },
                            'Outros Débitos': { name: "Status: Outros Débitos", color: "bg-amber-100 text-amber-700", border: "border-amber-200" },
                            'Estorno': { name: "Status: Estorno", color: "bg-red-100 text-red-700", border: "border-red-200" },
                            'Pendente': { name: "Status: Pendente", color: "bg-slate-100 text-slate-700", border: "border-slate-200" }
                          };
                          const label = statusLabels[status as keyof typeof statusLabels];
                          
                          return (
                            <div key={status} className={cn("p-6 rounded-3xl border bg-white shadow-sm glass-card", label.border)}>
                              <div className="flex items-center justify-between mb-4">
                                <span className={cn("px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest", label.color)}>
                                  {label.name}
                                </span>
                                <span className="text-navy font-black text-xl">{statusData.count} <span className="text-xs text-slate-400 uppercase">itens</span></span>
                              </div>
                              <div className="space-y-3">
                                <div className="flex justify-between text-xs font-bold">
                                  <span className="text-slate-400">VALOR CONTÁBIL</span>
                                  <span className="text-navy">{formatCurrency(statusData.contabil)}</span>
                                </div>
                                <div className="flex justify-between text-xs font-bold">
                                  <span className="text-slate-400">VALOR ICMS</span>
                                  <span className="text-navy">
                                    {formatCurrency(status === 'Outros Débitos' ? statusData.outros : statusData.icms)}
                                  </span>
                                </div>
                                {status === 'Outros Débitos' && (
                                  <div className="flex justify-between text-xs font-bold pt-2 border-t border-slate-100">
                                    <span className="text-amber-600">OUTROS DÉBITOS</span>
                                    <span className="text-amber-600">{formatCurrency(statusData.outros)}</span>
                                  </div>
                                )}
                                {status === 'Estorno' && (
                                  <div className="flex justify-between text-xs font-bold pt-2 border-t border-slate-100">
                                    <span className="text-red-600">ESTORNO DÉBITO</span>
                                    <span className="text-red-600">{formatCurrency(statusData.estorno)}</span>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>

                  {/* Integrated Triage Section (Not in PDF) */}
                  {pendingItems.length > 0 && (
                    <motion.div 
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="glass-card overflow-hidden border-2 border-amber-200"
                    >
                      <div className="p-8 bg-amber-50/50 border-b border-amber-100 flex justify-between items-center">
                        <div>
                          <h3 className="text-2xl font-black text-navy tracking-tighter uppercase">Itens Pendentes de Classificação</h3>
                          <p className="text-amber-700 text-xs font-bold uppercase tracking-widest mt-1">
                            Defina a situação para os {pendingItems.length} novos itens encontrados
                          </p>
                        </div>
                        <div className="flex gap-4">
                          <button 
                            onClick={saveTriageAndRefresh}
                            disabled={Object.keys(triageSelections).length === 0 || isSavingTriage}
                            className="px-8 py-4 bg-navy text-white rounded-2xl font-black text-sm flex items-center gap-3 hover:bg-navy-light shadow-xl transition-all disabled:opacity-50"
                          >
                            {isSavingTriage ? <RefreshCw className="animate-spin" size={20} /> : <CheckCircle2 size={20} />}
                            SALVAR E ATUALIZAR
                          </button>
                          <button 
                            onClick={() => pendingRulesFileInputRef.current?.click()}
                            disabled={isSavingTriage}
                            className="px-8 py-4 bg-white border-2 border-navy text-navy rounded-2xl font-black text-sm flex items-center gap-3 hover:bg-slate-50 transition-all disabled:opacity-50"
                          >
                            <Upload size={20} /> SUBIR PLANILHA
                          </button>
                        </div>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left">
                          <thead>
                            <tr className="bg-white/50 text-[10px] uppercase tracking-[0.2em] text-slate-400 font-black border-b border-amber-100">
                              <th className="px-8 py-5">NCM</th>
                              <th className="px-8 py-5">Item</th>
                              <th className="px-8 py-5">Condição ICMS</th>
                              <th className="px-8 py-5 text-center">Definir Situação</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-amber-100 bg-white/30">
                            {pendingItems.map((item) => {
                              const key = normalizeStr(item.item);
                              return (
                                <tr key={key} className="hover:bg-white/60 transition-colors">
                                  <td className="px-8 py-4">
                                    <div className="flex flex-col">
                                      <span className="font-mono text-[10px] font-black text-slate-400 tracking-tighter">{item.ncm || '-'}</span>
                                      <span className="text-[10px] font-bold text-slate-300 uppercase">{item.natureza}</span>
                                    </div>
                                  </td>
                                  <td className="px-8 py-4">
                                    <span className="text-sm font-black text-navy uppercase">{item.item}</span>
                                  </td>
                                  <td className="px-8 py-4 font-bold text-[10px] uppercase text-slate-400 tracking-widest">
                                    {item.hasIcms ? 'Com ICMS' : 'Sem ICMS'}
                                  </td>
                                  <td className="px-8 py-4">
                                    <div className="flex justify-center gap-2">
                                      {[
                                        { val: 1, label: 'Mantém', color: 'bg-emerald-100 text-emerald-700' },
                                        { val: 2, label: 'Outros Déb.', color: 'bg-amber-100 text-amber-700' },
                                        { val: 3, label: 'Estorno', color: 'bg-red-100 text-red-700' }
                                      ].map(opt => (
                                        <button
                                          key={opt.val}
                                          onClick={() => setTriageSelections(prev => ({ ...prev, [key]: opt.val as 1|2|3 }))}
                                          className={cn(
                                            "px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all",
                                            triageSelections[key] === opt.val 
                                              ? opt.color + " ring-2 ring-navy ring-offset-2" 
                                              : "bg-slate-50 text-slate-400 hover:bg-slate-100"
                                          )}
                                        >
                                          {opt.label}
                                        </button>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </motion.div>
                  )}

                  {/* Spreadsheet Table (Not in PDF) */}
                  <div className="glass-card overflow-hidden no-print">
                    <div className="p-8 border-b border-slate-100 flex flex-col md:flex-row justify-between items-center gap-6">
                      <div>
                        <h3 className="text-2xl font-black text-navy tracking-tighter uppercase">Relatório de Apuração Mensal</h3>
                        <p className="text-slate-400 text-xs font-bold uppercase tracking-widest mt-1">
                          Exibindo {filteredData.length} de {processedData.length} itens (Máx. 100)
                        </p>
                      </div>
                      <div className="relative w-full md:w-96">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-300" size={20} />
                        <input 
                          type="text" 
                          placeholder="Filtrar por item ou NCM..." 
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                          className="w-full pl-12 pr-6 py-4 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold focus:outline-none focus:border-navy transition-all"
                        />
                      </div>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50/80 text-[10px] uppercase tracking-[0.2em] text-slate-400 font-black border-b border-slate-100">
                            <th className="px-8 py-5">NCM</th>
                            <th className="px-8 py-5">Item</th>
                            <th className="px-8 py-5 text-right">Valor Contábil</th>
                            <th className="px-8 py-5 text-center">CST</th>
                            <th className="px-8 py-5 text-right">Valor ICMS</th>
                            <th className="px-8 py-5 text-right bg-amber-50/30 text-amber-600">Outros Débitos</th>
                            <th className="px-8 py-5 text-right bg-red-50/30 text-red-600">Estorno Débito</th>
                            <th className="px-8 py-5 text-center">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {filteredData.length > 0 ? filteredData.map((row) => (
                            <tr 
                              key={row.id} 
                              className={cn(
                                "transition-all group",
                                row.status === 'Outros Débitos' ? "bg-amber-50/40 hover:bg-amber-50/60" : 
                                row.status === 'Estorno' ? "bg-red-50/40 hover:bg-red-50/60" :
                                "hover:bg-slate-50/80"
                              )}
                            >
                              <td className="px-8 py-5 font-mono text-xs font-black text-slate-400 tracking-tighter">{row.ncm}</td>
                              <td className="px-8 py-5">
                                <div className="flex items-center gap-2">
                                  <span className="text-sm font-black text-navy uppercase">{row.item}</span>
                                </div>
                              </td>
                              <td className="px-8 py-5 text-right text-sm font-black text-slate-600">{formatCurrency(row.valorContabil)}</td>
                              <td className="px-8 py-5 text-center">
                                <span className="px-3 py-1 bg-white border border-slate-200 rounded-lg text-[10px] font-black text-slate-400">{row.cstIcms}</span>
                              </td>
                              <td className="px-8 py-5 text-right text-sm font-black text-blue-600">{formatCurrency(row.valorIcms)}</td>
                              <td className="px-8 py-5 text-right text-sm font-black text-amber-600">
                                {row.outrosDebitos > 0 ? formatCurrency(row.outrosDebitos) : '-'}
                              </td>
                              <td className="px-8 py-5 text-right text-sm font-black text-red-600">
                                {row.estornoDebito > 0 ? formatCurrency(row.estornoDebito) : '-'}
                              </td>
                              <td className="px-8 py-5 text-center">
                                <span className={cn(
                                  "px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest",
                                  row.status === 'Normal' ? "bg-emerald-100 text-emerald-600" : 
                                  row.status === 'Outros Débitos' ? "bg-amber-100 text-amber-600" :
                                  row.status === 'Estorno' ? "bg-red-100 text-red-600" :
                                  "bg-slate-100 text-slate-600"
                                )}>
                                  {row.status}
                                </span>
                              </td>
                            </tr>
                          )) : (
                            <tr>
                              <td colSpan={8} className="px-8 py-20 text-center">
                                <div className="flex flex-col items-center gap-4">
                                  <FileSpreadsheet size={48} className="text-slate-200" />
                                  <p className="text-slate-400 font-black uppercase text-xs tracking-widest">Nenhum dado processado.</p>
                                </div>
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </motion.div>
              )}
            </motion.div>
          )}

          {activeTab === 'expansao' && (
            <motion.div 
              key="expansao"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.05 }}
              className="flex flex-col items-center justify-center py-24"
            >
              <div className="text-center mb-16">
                <h2 className="text-5xl font-black text-navy mb-6 tracking-tight">Expansão de Base</h2>
                <p className="text-slate-400 text-xl max-w-2xl mx-auto font-medium">
                  Alimente o banco de regras com novos itens. O sistema <span className="text-navy font-bold">adicionará</span> as novas regras às existentes.
                </p>
                
                <div className="mt-8 p-6 bg-blue-50 border border-blue-100 rounded-3xl max-w-md mx-auto">
                  <h4 className="text-blue-800 font-black text-xs uppercase tracking-widest mb-3 flex items-center justify-center gap-2">
                    <FileSpreadsheet size={16} /> Formato da Planilha
                  </h4>
                  <div className="flex flex-wrap justify-center gap-2">
                    {['NCM', 'NATUREZA', 'ITEM', 'BASE CÁLCULO ICM', 'SITUAÇÃO'].map(col => (
                      <span key={col} className="px-2 py-1 bg-white border border-blue-200 rounded text-[10px] font-black text-blue-600">
                        {col}
                      </span>
                    ))}
                  </div>
                  <p className="text-[10px] text-blue-600/70 font-bold mt-3 uppercase tracking-tight">
                    A coluna <span className="text-blue-800 underline">VALOR ICMS</span> deve conter <span className="text-blue-800">SIM</span> ou <span className="text-blue-800">NÃO</span>.
                  </p>
                  <p className="text-[10px] text-blue-600/50 font-medium mt-1 uppercase tracking-tight">
                    SITUAÇÃO: 1 (Normal), 2 (Outros Débitos 20,5%), 3 (Estorno ICMS).
                  </p>
                </div>
              </div>

              <div 
                onClick={() => {
                  setRuleImportMode('append');
                  rulesFileInputRef.current?.click();
                }}
                className="w-full max-w-2xl group cursor-pointer"
              >
                <div className="relative p-12 border-4 border-dashed rounded-[40px] bg-white border-slate-200 hover:border-navy hover:shadow-2xl transition-all duration-500 flex flex-col items-center">
                  <div className="w-24 h-24 rounded-3xl bg-slate-50 flex items-center justify-center mb-8 group-hover:bg-navy group-hover:rotate-6 transition-all duration-500">
                    <Plus className="text-slate-300 group-hover:text-white" size={48} />
                  </div>
                  <h3 className="text-2xl font-black text-navy mb-2">Importar Novas Regras</h3>
                  <p className="text-slate-400 font-bold tracking-wide">Clique para selecionar a planilha de expansão</p>
                  
                  <div className="absolute -bottom-6 left-1/2 -translate-x-1/2 bg-navy text-white px-8 py-3 rounded-full font-black text-sm shadow-xl group-hover:scale-110 transition-transform uppercase">
                    Expandir Banco
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'base' && (
            <motion.div 
              key="base"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              className="max-w-5xl mx-auto space-y-8"
            >
              <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-6">
                <div>
                  <h2 className="text-4xl font-black text-navy tracking-tighter uppercase">Banco de Regras</h2>
                  <p className="text-slate-400 font-bold uppercase tracking-widest text-xs mt-2">Mapeamento Exato de Itens e Situações</p>
                </div>
                <div className="flex items-center gap-6">
                  <button 
                    onClick={exportRules}
                    className="px-6 py-4 bg-white border-2 border-navy text-navy rounded-2xl font-black text-sm flex items-center gap-3 hover:bg-slate-50 transition-all"
                  >
                    <Download size={20} /> BAIXAR BANCO (EXCEL)
                  </button>
                  <button 
                    onClick={() => {
                      if (confirm("ATENÇÃO: Isso irá APAGAR todas as regras atuais e substituir pelas da planilha. Continuar?")) {
                        setRuleImportMode('replace');
                        rulesFileInputRef.current?.click();
                      }
                    }}
                    disabled={isImportingRules}
                    className="px-6 py-4 bg-navy text-white rounded-2xl font-black text-sm flex items-center gap-3 hover:bg-navy-light shadow-xl transition-all disabled:opacity-50"
                  >
                    {isImportingRules ? <RefreshCw className="animate-spin" size={20} /> : <RefreshCw size={20} />}
                    SUBSTITUIR BANCO NOVO (EXCEL)
                  </button>
                </div>
              </div>

              {/* Legend of Rules */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="p-6 bg-white border border-slate-100 rounded-3xl shadow-sm">
                  <div className="flex items-center gap-3 mb-3">
                    <div className="w-8 h-8 bg-emerald-100 rounded-lg flex items-center justify-center font-black text-emerald-600">✓</div>
                    <h4 className="font-black text-navy text-xs uppercase tracking-widest">Item Normal</h4>
                  </div>
                  <p className="text-xs text-slate-400 font-bold">O item possui ICMS destacado ou não está na lista de exigência.</p>
                </div>
                <div className="p-6 bg-amber-50 border border-amber-100 rounded-3xl shadow-sm">
                  <div className="flex items-center gap-3 mb-3">
                    <div className="w-8 h-8 bg-amber-500 rounded-lg flex items-center justify-center font-black text-white">!</div>
                    <h4 className="font-black text-amber-700 text-xs uppercase tracking-widest">Outros Débitos</h4>
                  </div>
                  <p className="text-xs text-amber-600/70 font-bold">Se o item estiver na lista de regras e a condição de ICMS for atendida, será calculado 100% de Outros Débitos.</p>
                </div>
                <div className="p-6 bg-red-50 border border-red-100 rounded-3xl shadow-sm">
                  <div className="flex items-center gap-3 mb-3">
                    <div className="w-8 h-8 bg-red-500 rounded-lg flex items-center justify-center font-black text-white">↺</div>
                    <h4 className="font-black text-red-700 text-xs uppercase tracking-widest">Estorno de Débito</h4>
                  </div>
                  <p className="text-xs text-red-600/70 font-bold">Se o item estiver na lista de regras e a condição de ICMS for atendida, será realizado o estorno de 100% do valor.</p>
                </div>
              </div>

              <div className="glass-card overflow-hidden">
                <div className="p-8 border-b border-slate-100 flex flex-col md:flex-row justify-between items-center gap-4">
                  <div>
                    <h3 className="text-xl font-black text-navy uppercase tracking-tighter">Itens na Base de Regras ({rules.length})</h3>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mt-1">Exibindo apenas os 100 primeiros itens para manter o desempenho</p>
                  </div>
                  <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-4 py-2 bg-slate-100 rounded-lg">Lista Resumida</div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead>
                      <tr className="bg-slate-50 text-[10px] font-black uppercase tracking-widest text-slate-400 border-b border-slate-100">
                        <th className="px-8 py-5">NCM</th>
                        <th className="px-8 py-5">Natureza</th>
                        <th className="px-8 py-5">Nome / Descrição</th>
                        <th className="px-8 py-5">Condição ICMS</th>
                        <th className="px-8 py-5">Situação</th>
                        <th className="px-8 py-5 text-right">Ações</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {rules.slice(0, 100).map((rule, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                        <td className="px-8 py-5 font-mono text-[10px] font-black text-slate-400">{rule.ncm}</td>
                        <td className="px-8 py-5">
                          <span className={cn(
                            "text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-lg",
                            rule.situacao === 2 ? "text-amber-600 bg-amber-100" : 
                            rule.situacao === 3 ? "text-red-600 bg-red-100" : 
                            "text-slate-600 bg-slate-100"
                          )}>
                            {rule.natureza || rule.acao}
                          </span>
                        </td>
                        <td className="px-8 py-5 font-black text-navy uppercase tracking-tight">{rule.item}</td>
                        <td className="px-8 py-5">
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                            {rule.hasIcms ? 'Com ICMS' : 'Sem ICMS'}
                          </span>
                        </td>
                        <td className="px-8 py-5">
                          <div className="flex items-center gap-2">
                            <span className="w-6 h-6 rounded-full bg-navy text-white flex items-center justify-center text-[10px] font-black">
                              {rule.situacao}
                            </span>
                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                              {rule.situacao === 1 ? 'Mantém' : rule.situacao === 2 ? 'Outros Déb.' : 'Estorno'}
                            </span>
                          </div>
                        </td>
                        <td className="px-8 py-5 text-right">
                          <button 
                            onClick={() => removeRule(rule.item)}
                            className="p-2 text-slate-300 hover:text-red-500 transition-colors"
                          >
                            <Trash2 size={18} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  </table>
                </div>
              </div>
            </motion.div>
        )}
        </AnimatePresence>
      </main>

      <footer className="p-10 text-center">
        <div className="flex justify-center items-center gap-4 mb-4">
          <div className="h-px w-12 bg-slate-200"></div>
          <p className="text-[10px] font-black text-slate-300 uppercase tracking-[0.4em]">ANÁLISE ICMS CP Database Engine</p>
          <div className="h-px w-12 bg-slate-200"></div>
        </div>
      </footer>
    </div>
  );
}
