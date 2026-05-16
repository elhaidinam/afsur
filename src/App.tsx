/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import * as d3 from 'd3';
import {
  FileUp,
  LayoutDashboard,
  BarChart3,
  Trees,
  Languages,
  Database,
  Search,
  Filter,
  Activity,
  Info,
  ChevronRight,
  Loader2,
  Globe,
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Line,
  ComposedChart,
  Cell,
  Legend,
  Scatter,
  ErrorBar,
} from 'recharts';
import { GoogleGenAI, Type } from '@google/genai';
import { cn, safeNumeric, cleanName } from './lib/utils';

// --- Types ---
type Language = 'EN' | 'FR' | 'PT';

const TRANSLATIONS = {
  EN: {
    appName: "AFSUR26 - Safety Updates",
    sidebar: {
      dashboards: "Dashboards",
      overview: "ICSR Overview",
      serious: "Serious ICSRs",
      criteria: "Search Criteria",
      disprop: "Disproportionality",
      subgroup: "Subgroup Analysis",
      status: "Status",
      ready: "Data Ready",
      awaiting: "Awaiting Upload",
      upload: "Upload XLSX"
    },
    tabs: {
      overview: "📈 ICSR Overview",
      disprop: "📊 IC / IC025",
      forest: "🌲 Forest Plot",
      ai: "🧠 AI Support"
    },
    dashboard: {
      geoDist: "Geographic Distribution",
      breakdown: "Breakdown of reported cases by dimension.",
      ratePer: "Rate per 1M",
      signalTable: "Signal Detection Table",
      activeSignals: "Active Signals",
      clinicalNotes: "Clinical Notes",
      detected: "Signal Detected"
    },
    ai: {
      title: "Clinical Synonym Grouping",
      subtitle: "Identifying synonymous MedDRA terms to streamline signal evaluation.",
      button: "Identify Synonym Clusters",
      loading: "Analyzing Medical groups..."
    }
  },
  FR: {
    appName: "AFSUR26 - Safety Updates",
    sidebar: {
      dashboards: "Tableaux de bord",
      overview: "Aperçu ICSR",
      serious: "ICSR Graves",
      criteria: "Critères de recherche",
      disprop: "Disproportionnalité",
      subgroup: "Analyse par sous-groupe",
      status: "Statut",
      ready: "Données prêtes",
      awaiting: "En attente d'import",
      upload: "Charger XLSX"
    },
    tabs: {
      overview: "📈 Aperçu ICSR",
      disprop: "📊 IC / IC025",
      forest: "🌲 Forest Plot",
      ai: "🧠 Support IA"
    },
    dashboard: {
      geoDist: "Distribution Géographique",
      breakdown: "Répartition des cas signalés par dimension.",
      ratePer: "Taux par 1M",
      signalTable: "Tableau de détection des signaux",
      activeSignals: "Signaux Actifs",
      clinicalNotes: "Notes cliniques",
      detected: "Signal Détecté"
    },
    ai: {
      title: "Groupement de synonymes cliniques",
      subtitle: "Identification des termes MedDRA synonymes pour rationaliser l'évaluation des signaux.",
      button: "Identifier les groupes de synonymes",
      loading: "Analyse des groupes médicaux..."
    }
  },
  PT: {
    appName: "AFSUR26 - Safety Updates",
    sidebar: {
      dashboards: "Painéis",
      overview: "Visão Geral ICSR",
      serious: "ICSRs Graves",
      criteria: "Critérios de Pesquisa",
      disprop: "Desproporcionalidade",
      subgroup: "Análise de Subgrupo",
      status: "Status",
      ready: "Dados Prontos",
      awaiting: "Aguardando Upload",
      upload: "Carregar XLSX"
    },
    tabs: {
      overview: "📈 Visão Geral ICSR",
      disprop: "📊 IC / IC025",
      forest: "🌲 Forest Plot",
      ai: "🧠 Suporte de IA"
    },
    dashboard: {
      geoDist: "Distribuição Geográfica",
      breakdown: "Detalhamento dos casos notificados por dimensão.",
      ratePer: "Taxa por 1M",
      signalTable: "Tabela de detecção de sinal",
      activeSignals: "Sinais Ativos",
      clinicalNotes: "Notas Clínicas",
      detected: "Sinal Detectado"
    },
    ai: {
      title: "Agrupamento de Sinônimos Clínicos",
      subtitle: "Identificação de termos MedDRA sinônimos para agilizar a avaliação do sinal.",
      button: "Identificar Clusters de Sinônimos",
      loading: "Analisando grupos médicos..."
    }
  }
};

interface RawData {
  [key: string]: any;
}

interface OverviewItem {
  dimension: string;
  total: number;
  serious: number;
  nonSerious: number;
  rowDenominator?: number;
  frequency?: number;
}

interface RawOverviewData {
  dimension: string;
  count: number;
  rowDenominator?: number;
}

interface DispropItem {
  reaction: string;
  ic: number;
  ic025: number;
  nObserved: number;
  nSerious: number;
  isSignal: boolean;
}

interface SubgroupItem {
  group: string;
  ic: number;
  ic0005: number;
  ic9995: number;
  isSignal: boolean;
}

interface SynonymGroup {
  groupName: string;
  terms: string[];
  clinicalContext: string;
}

// --- Helper Components ---

function WorldMap({ data, lang }: { data: OverviewItem[], lang: Language }) {
  const [geoData, setGeoData] = useState<any>(null);
  const t = TRANSLATIONS[lang];
  
  const svgRef = useCallback((node: SVGSVGElement | null) => {
    if (!node || !geoData) return;

    const width = 800;
    const height = 400;

    const svg = d3.select(node);
    svg.selectAll("*").remove();

    const projection = d3.geoMercator()
      .scale(350)
      .center([20, 5])
      .translate([width / 2, height / 2]);

    const path = d3.geoPath().projection(projection);

    const colorScale = d3.scaleThreshold<number, string>()
      .domain([10, 100, 500, 1000, 5000])
      .range(["#f2f2f2", "#d1e5f0", "#a6cee3", "#67a9cf", "#3288bd", "#1E7FB8"]);

    // Create a lookup for quick mapping
    const dataMap = new Map(data.map(d => [d.dimension.toLowerCase().trim(), d.total]));

    svg.append("g")
      .selectAll("path")
      .data(geoData.features)
      .enter()
      .append("path")
      .attr("d", path as any)
      .attr("fill", (d: any) => {
        const name = d.properties.name.toLowerCase().trim();
        const value = dataMap.get(name) || 0;
        return colorScale(value);
      })
      .attr("stroke", "#ffffff")
      .attr("stroke-width", 0.5)
      .append("title")
      .text((d: any) => {
        const name = d.properties.name;
        const value = dataMap.get(name.toLowerCase().trim()) || 0;
        const reportsLabel = lang === 'FR' ? 'signalements' : (lang === 'PT' ? 'notificações' : 'reports');
        return `${name}: ${value} ${reportsLabel}`;
      });
  }, [geoData, data, lang]);

  useEffect(() => {
    fetch('https://raw.githubusercontent.com/holtzy/D3-graph-gallery/master/DATA/world.geojson')
      .then(res => res.json())
      .then(setGeoData)
      .catch(err => console.error("Map load error:", err));
  }, []);

  return (
    <div className="w-full bg-white p-6 rounded-3xl border border-[#141414]/10 shadow-sm mt-6">
      <div className="flex items-center gap-2 mb-4">
        <Globe size={18} className="text-[#1E7FB8]" />
        <h3 className="font-bold tracking-tight">{t.dashboard.geoDist}</h3>
      </div>
      <div className="w-full overflow-hidden flex justify-center bg-[#FDFCFB]/50 rounded-2xl">
        <svg
          viewBox="0 0 800 400"
          ref={svgRef}
          className="w-full h-auto max-w-[1000px]"
        />
      </div>
      <div className="mt-4 flex flex-wrap gap-4 text-[10px] font-bold text-[#141414]/40 uppercase tracking-widest">
        <div className="flex items-center gap-1"><div className="w-3 h-3 bg-[#f2f2f2] rounded" /> 0</div>
        <div className="flex items-center gap-1"><div className="w-3 h-3 bg-[#d1e5f0] rounded" /> 1-10</div>
        <div className="flex items-center gap-1"><div className="w-3 h-3 bg-[#a6cee3] rounded" /> 11-100</div>
        <div className="flex items-center gap-1"><div className="w-3 h-3 bg-[#67a9cf] rounded" /> 101-500</div>
        <div className="flex items-center gap-1"><div className="w-3 h-3 bg-[#3288bd] rounded" /> 501-5000</div>
        <div className="flex items-center gap-1"><div className="w-3 h-3 bg-[#1E7FB8] rounded" /> 5000+</div>
      </div>
    </div>
  );
}

function GroupCard(group: SynonymGroup, index: number, lang: Language) {
  const t = TRANSLATIONS[lang];
  return (
    <motion.div
      key={index}
      initial={{ opacity: 0, scale: 0.95 }}
      animate={{ opacity: 1, scale: 1 }}
      transition={{ delay: index * 0.05 }}
      className="bg-white p-6 rounded-3xl border border-[#141414]/10 shadow-sm hover:border-[#1E7FB8]/30 transition-all group"
    >
      <div className="flex items-center gap-3 mb-4">
        <div className="p-2 bg-[#1E7FB8]/10 rounded-lg group-hover:bg-[#1E7FB8]/20 transition-colors">
          <Database size={16} className="text-[#1E7FB8]" />
        </div>
        <h3 className="font-bold tracking-tight text-lg">{group.groupName}</h3>
      </div>
      
      <div className="space-y-4">
        <div className="flex flex-wrap gap-2">
          {group.terms.map((term, i) => (
            <span key={i} className="px-3 py-1 bg-[#141414]/5 text-[#141414]/70 text-[10px] font-bold rounded-full uppercase tracking-wider">
              {term}
            </span>
          ))}
        </div>
        
        <div className="pt-4 border-t border-[#141414]/5">
          <p className="text-[10px] font-bold text-[#1E7FB8] uppercase tracking-widest mb-1 flex items-center gap-1">
            <Info size={10} /> {t.dashboard.clinicalNotes}
          </p>
          <p className="text-xs text-[#141414]/60 leading-relaxed italic">
            {group.clinicalContext}
          </p>
        </div>
      </div>
    </motion.div>
  );
}

// --- App Component ---
export default function App() {
  const [activeTab, setActiveTab] = useState<'icsr' | 'disprop' | 'forest' | 'synonyms'>('icsr');
  const [language, setLanguage] = useState<Language>('EN');
  const t = TRANSLATIONS[language];

  const logoSrc = useMemo(() => {
    switch (language) {
      case 'FR': return 'afrofr.png';
      case 'PT': return 'afropt.png';
      default: return 'afroen.png';
    }
  }, [language]);
  
  // File states
  const [overviewFile, setOverviewFile] = useState<File | null>(null);
  const [overviewSeriousFile, setOverviewSeriousFile] = useState<File | null>(null);
  const [dispropFile, setDispropFile] = useState<File | null>(null);
  const [subgroupFile, setSubgroupFile] = useState<File | null>(null);
  
  // Sheet states
  const [overviewSheets, setOverviewSheets] = useState<string[]>([]);
  const [subgroupSheets, setSubgroupSheets] = useState<string[]>([]);
  
  const [selectedOverviewSheet, setSelectedOverviewSheet] = useState<string>('');
  const [selectedSubgroupSheet, setSelectedSubgroupSheet] = useState<string>('');

  // Parsed Raw Datas
  const [rawTotalData, setRawTotalData] = useState<RawOverviewData[]>([]);
  const [rawSeriousData, setRawSeriousData] = useState<RawOverviewData[]>([]);
  
  // Input states
  const [denominator, setDenominator] = useState<number | ''>('');
  const [denomType, setDenomType] = useState<string>('Doses');
  
  // Final Data states
  const [overviewData, setOverviewData] = useState<OverviewItem[]>([]);
  const [dispropData, setDispropData] = useState<DispropItem[]>([]);
  const [subgroupData, setSubgroupData] = useState<SubgroupItem[]>([]);
  const [subgroupMeta, setSubgroupMeta] = useState<string>('');
  const [isLoading, setIsLoading] = useState(false);

  // AI Synonyms states
  const [synonymGroups, setSynonymGroups] = useState<SynonymGroup[]>([]);
  const [isAiLoading, setIsAiLoading] = useState(false);

  // --- Excel Parsing Logic ---

  const handleFileUpload = async (file: File, type: 'overview' | 'overview_serious' | 'disprop' | 'subgroup') => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetNames = workbook.SheetNames;

      if (type === 'overview') {
        setOverviewFile(file);
        setOverviewSheets(sheetNames);
        setSelectedOverviewSheet(sheetNames[0]);
      } else if (type === 'overview_serious') {
        setOverviewSeriousFile(file);
      } else if (type === 'disprop') {
        const sheet = sheetNames.find(s => s === 'Combination list') || sheetNames[0];
        const worksheet = workbook.Sheets[sheet];
        const json = XLSX.utils.sheet_to_json(worksheet) as RawData[];
        
        const processed = json.map(row => {
          const keys = Object.keys(row);
          // Prioritize MedDRA specific naming
          const reactionKey = keys.find(k => 
            cleanName(k) === 'Reaction (PT)' || 
            cleanName(k) === 'Reported preferred terms (MedDRA)'
          ) || keys[0];
          const ic = safeNumeric(row['IC']);
          const ic025 = safeNumeric(row['IC025']);
          return {
            reaction: String(row[reactionKey] || 'Unknown'),
            ic,
            ic025,
            nObserved: safeNumeric(row['Nobserved']),
            nSerious: safeNumeric(row['Nserious']),
            isSignal: ic025 > 0
          };
        });
        setDispropData(processed);
        setDispropFile(file);
      } else if (type === 'subgroup') {
        setSubgroupFile(file);
        const validSheets = sheetNames.filter(s => s.toLowerCase() !== 'search criteria');
        setSubgroupSheets(validSheets);
        setSelectedSubgroupSheet(validSheets[0]);
        
        // Extract meta
        const metaSheet = sheetNames.find(s => s.toLowerCase() === 'search criteria');
        if (metaSheet) {
          const mWs = workbook.Sheets[metaSheet];
          const mJson = XLSX.utils.sheet_to_json(mWs) as RawData[];
          if (mJson.length > 0) {
            const metaStr = Object.entries(mJson[0])
              .map(([k, v]) => `${k}: ${v}`)
              .join(' | ');
            setSubgroupMeta(metaStr);
          }
        }
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Re-parse TOTAL overview
  useEffect(() => {
    if (!overviewFile || !selectedOverviewSheet) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[selectedOverviewSheet];
      if (!worksheet) {
        console.error("Sheet not found in Total file:", selectedOverviewSheet);
        return;
      }
      const json = XLSX.utils.sheet_to_json(worksheet) as any[];
      
      const processed: RawOverviewData[] = json.map(row => {
          const keys = Object.keys(row);
          const denomKey = keys.find(k => k.toLowerCase() === 'denominator');
          return {
            dimension: String(row[keys[0]] || 'Unknown'),
            count: safeNumeric(row[keys[1]]),
            rowDenominator: denomKey ? safeNumeric(row[denomKey]) : undefined
          };
        }).filter(item => !isNaN(item.count));
        
      setRawTotalData(processed);
    };
    reader.readAsArrayBuffer(overviewFile);
  }, [overviewFile, selectedOverviewSheet]);

  // Re-parse SERIOUS overview
  useEffect(() => {
    if (!overviewSeriousFile || !selectedOverviewSheet) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[selectedOverviewSheet];
      if (!worksheet) {
        console.error("Sheet not found in Serious file:", selectedOverviewSheet);
        return;
      }
      const json = XLSX.utils.sheet_to_json(worksheet) as any[];
      
      const processed: RawOverviewData[] = json.map(row => {
          const keys = Object.keys(row);
          return {
            dimension: String(row[keys[0]] || 'Unknown'),
            count: safeNumeric(row[keys[1]])
          };
        }).filter(item => !isNaN(item.count));
        
      setRawSeriousData(processed);
    };
    reader.readAsArrayBuffer(overviewSeriousFile);
  }, [overviewSeriousFile, selectedOverviewSheet]);

  // Combine raw data into OverviewItems
  useEffect(() => {
    if (!rawTotalData || rawTotalData.length === 0) {
        setOverviewData([]);
        return;
    }

    try {
      const seriousMap = new Map<string, number>(rawSeriousData ? rawSeriousData.map(d => [d.dimension, d.count]) : []);
      
      const combined: OverviewItem[] = rawTotalData.map(total => {
          const serious = seriousMap.get(total.dimension) || 0;
          const totalCount = total.count || 0;
          const nonSerious = Math.max(0, totalCount - serious);
          return {
              dimension: total.dimension || 'Unknown',
              total: totalCount,
              serious,
              nonSerious,
              rowDenominator: total.rowDenominator
          };
      });

      setOverviewData(combined);
    } catch (error) {
      console.error("Error combining overview data:", error);
      setOverviewData([]);
    }
  }, [rawTotalData, rawSeriousData]);

  // Re-parse subgroup when sheet changes
  useEffect(() => {
    if (!subgroupFile || !selectedSubgroupSheet) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[selectedSubgroupSheet];
      if (!worksheet) return;
      const json = XLSX.utils.sheet_to_json(worksheet) as RawData[];
      
      const processed: SubgroupItem[] = json.map(row => {
        const keys = Object.keys(row);
        const groupKey = keys[0];
        const ic0005 = safeNumeric(row['IC0005']);
        return {
          group: String(row[groupKey]),
          ic: safeNumeric(row['IC']),
          ic0005,
          ic9995: safeNumeric(row['IC9995']),
          isSignal: ic0005 > 0
        };
      }).filter(item => !isNaN(item.ic));
      
      setSubgroupData(processed);
    };
    reader.readAsArrayBuffer(subgroupFile);
  }, [subgroupFile, selectedSubgroupSheet]);

  // --- AI Logic for Grouping Synonyms ---
  const fetchSynonyms = async () => {
    if (dispropData.length === 0 || isAiLoading) return;
    
    setIsAiLoading(true);
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      // Use more PTs for grouping
      const pts = dispropData.slice(0, 30).map(d => d.reaction);
      
      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: `Given the following Pharmacovigilance Preferred Terms (MedDRA): [${pts.join(', ')}]. 
        Group together the terms that are medical synonyms or clinically related into logical clusters. 
        Each group should represent a unique clinical concept. 
        Return as a JSON array of objects with properties: groupName, terms (array of original PTs), clinicalContext.`,
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                groupName: { type: Type.STRING },
                terms: { 
                  type: Type.ARRAY,
                  items: { type: Type.STRING }
                },
                clinicalContext: { type: Type.STRING }
              },
              required: ['groupName', 'terms', 'clinicalContext']
            }
          }
        }
      });
      
      const data = JSON.parse(response.text || '[]') as SynonymGroup[];
      setSynonymGroups(data);
    } catch (error) {
      console.error('AI Error:', error);
    } finally {
      setIsAiLoading(false);
    }
  };

  // --- Computed Data ---
  const overviewChartData = useMemo(() => {
    const sorted = [...overviewData].sort((a, b) => b.total - a.total).slice(0, 20);
    const globalDenom = Number(denominator);
    
    return sorted.map(d => {
      const activeDenom = d.rowDenominator || globalDenom;
      return {
        ...d,
        frequency: (activeDenom && activeDenom > 0) ? (d.total / activeDenom) * 1000000 : undefined
      };
    });
  }, [overviewData, denominator]);

  const dispropChartData = useMemo(() => {
    return [...dispropData].sort((a, b) => b.ic - a.ic).slice(0, 30);
  }, [dispropData]);

  // --- Render Helpers ---

  return (
    <div className="flex h-screen bg-[#FDFCFB] text-[#141414] font-sans selection:bg-[#1E7FB8]/20">
      {/* Sidebar */}
      <aside className="w-80 border-r border-[#141414]/10 bg-[#E4E3E0] flex flex-col p-6 space-y-8 overflow-y-auto">
        <div className="space-y-4">
          <div className="flex gap-2 p-1 bg-white/50 rounded-xl border border-[#141414]/5">
            {(['EN', 'FR', 'PT'] as Language[]).map(l => (
              <button
                key={l}
                onClick={() => setLanguage(l)}
                className={cn(
                  "flex-1 py-1.5 text-[10px] font-bold rounded-lg transition-all",
                  language === l ? "bg-[#1E7FB8] text-white shadow-md shadow-[#1E7FB8]/20" : "text-[#141414]/40 hover:bg-white"
                )}
              >
                {l}
              </button>
            ))}
          </div>

          <div className="flex flex-col items-center gap-4 py-2 border-b border-[#141414]/10">
            <img src={logoSrc} alt="WHO Logo" className="h-24 w-auto object-contain" />
            <div className="text-center">
              <h1 className="text-xl font-black tracking-tighter uppercase leading-none">{t.appName}</h1>
              <p className="text-[10px] font-bold text-[#1E7FB8] uppercase tracking-widest mt-1">Intelligence System</p>
            </div>
          </div>
        </div>

        <div className="space-y-6">
          <div className="space-y-3">
            <label className="text-[11px] font-bold uppercase tracking-widest opacity-60 flex items-center gap-2">
              <LayoutDashboard size={14} /> {t.sidebar.dashboards}
            </label>
            <label className="text-[10px] font-bold uppercase tracking-widest opacity-40 block px-1">
               {t.sidebar.overview}
            </label>
            <div className="relative group">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0], 'overview')}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
              />
              <div className="p-3 border-2 border-dashed border-[#141414]/20 rounded-xl group-hover:border-[#1E7FB8] transition-colors flex items-center justify-center gap-2 bg-white/50">
                <FileUp size={18} className="text-[#1E7FB8]" />
                <span className="text-sm font-medium truncate">
                  {overviewFile ? overviewFile.name : 'Vigilyze overview.xlsx'}
                </span>
              </div>
            </div>
            <label className="text-[11px] font-bold uppercase tracking-widest opacity-60 flex items-center gap-2 pt-2">
              <LayoutDashboard size={14} /> {t.sidebar.serious}
            </label>
            <div className="relative group">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0], 'overview_serious')}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
              />
              <div className="p-3 border-2 border-dashed border-[#141414]/20 rounded-xl group-hover:border-[#1E7FB8] transition-colors flex items-center justify-center gap-2 bg-white/50">
                <FileUp size={18} className="text-[#1E7FB8]" />
                <span className="text-sm font-medium truncate">
                  {overviewSeriousFile ? overviewSeriousFile.name : 'Vigilyze overview serious.xlsx'}
                </span>
              </div>
            </div>

            {overviewSheets.length > 0 && (
              <div className="space-y-2">
                <label className="text-[10px] font-bold uppercase tracking-widest opacity-40 block px-1">
                  {t.sidebar.criteria}
                </label>
                <select
                  value={selectedOverviewSheet}
                  onChange={(e) => setSelectedOverviewSheet(e.target.value)}
                  className="w-full p-2 bg-white border border-[#141414]/10 rounded-lg text-xs focus:ring-2 focus:ring-[#1E7FB8] outline-none"
                >
                  {overviewSheets.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            )}

            <label className="text-[11px] font-bold uppercase tracking-widest opacity-60 flex items-center gap-2 pt-2">
              <Filter size={14} /> {t.sidebar.disprop}
            </label>
            <div className="relative group">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0], 'disprop')}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
              />
              <div className="p-3 border-2 border-dashed border-[#141414]/20 rounded-xl group-hover:border-[#1E7FB8] transition-colors flex items-center justify-center gap-2 bg-white/50">
                <FileUp size={18} className="text-[#1E7FB8]" />
                <span className="text-sm font-medium truncate">
                  {dispropFile ? dispropFile.name : t.sidebar.upload}
                </span>
              </div>
            </div>
          </div>

          <div className="space-y-3">
            <label className="text-[11px] font-bold uppercase tracking-widest opacity-60 flex items-center gap-2">
              <Activity size={14} /> {t.sidebar.subgroup}
            </label>
            <div className="relative group">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0], 'subgroup')}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
              />
              <div className="p-3 border-2 border-dashed border-[#141414]/20 rounded-xl group-hover:border-[#1E7FB8] transition-colors flex items-center justify-center gap-2 bg-white/50">
                <FileUp size={18} className="text-[#1E7FB8]" />
                <span className="text-sm font-medium truncate">
                  {subgroupFile ? subgroupFile.name : t.sidebar.upload}
                </span>
              </div>
            </div>
            {subgroupSheets.length > 0 && (
              <select
                value={selectedSubgroupSheet}
                onChange={(e) => setSelectedSubgroupSheet(e.target.value)}
                className="w-full p-2.5 bg-white border border-[#141414]/10 rounded-lg text-sm focus:ring-2 focus:ring-[#1E7FB8] outline-none"
              >
                {subgroupSheets.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            )}
          </div>
        </div>

        <div className="mt-auto pt-6 border-t border-[#141414]/10">
          <div className="p-4 bg-[#1E7FB8]/5 rounded-xl border border-[#1E7FB8]/20">
            <p className="text-[10px] text-[#1E7FB8] font-bold uppercase tracking-wider mb-1">{t.sidebar.status}</p>
            <div className="flex items-center gap-2 text-xs font-medium">
              <div className={cn("w-2 h-2 rounded-full animate-pulse", overviewFile ? "bg-green-500" : "bg-[#141414]/20")} />
              {overviewFile ? t.sidebar.ready : t.sidebar.awaiting}
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {/* Header Tabs */}
        <header className="h-20 border-bottom border-[#141414]/10 bg-white flex items-center justify-between px-10 shrink-0">
          <nav className="flex gap-8">
            <button
              onClick={() => setActiveTab('icsr')}
              className={cn(
                "relative h-20 flex items-center text-sm font-bold tracking-tight transition-colors",
                activeTab === 'icsr' ? "text-[#1E7FB8]" : "text-[#141414]/40 hover:text-[#141414]"
              )}
            >
              {t.tabs.overview}
              {activeTab === 'icsr' && <motion.div layoutId="tab" className="absolute bottom-0 left-0 right-0 h-1 bg-[#1E7FB8]" />}
            </button>
            <button
              onClick={() => setActiveTab('disprop')}
              className={cn(
                "relative h-20 flex items-center text-sm font-bold tracking-tight transition-colors",
                activeTab === 'disprop' ? "text-[#1E7FB8]" : "text-[#141414]/40 hover:text-[#141414]"
              )}
            >
              {t.tabs.disprop}
              {activeTab === 'disprop' && <motion.div layoutId="tab" className="absolute bottom-0 left-0 right-0 h-1 bg-[#1E7FB8]" />}
            </button>
            <button
              onClick={() => setActiveTab('forest')}
              className={cn(
                "relative h-20 flex items-center text-sm font-bold tracking-tight transition-colors",
                activeTab === 'forest' ? "text-[#1E7FB8]" : "text-[#141414]/40 hover:text-[#141414]"
              )}
            >
              {t.tabs.forest}
              {activeTab === 'forest' && <motion.div layoutId="tab" className="absolute bottom-0 left-0 right-0 h-1 bg-[#1E7FB8]" />}
            </button>
            <button
              onClick={() => setActiveTab('synonyms')}
              className={cn(
                "relative h-20 flex items-center text-sm font-bold tracking-tight transition-colors",
                activeTab === 'synonyms' ? "text-[#1E7FB8]" : "text-[#141414]/40 hover:text-[#141414]"
              )}
            >
              {t.tabs.ai}
              {activeTab === 'synonyms' && <motion.div layoutId="tab" className="absolute bottom-0 left-0 right-0 h-1 bg-[#1E7FB8]" />}
            </button>
          </nav>
          
          <div className="flex items-center gap-4">
            <button className="p-2 text-[#141414]/40 hover:text-[#141414] transition-colors"><Search size={20} /></button>
            <button className="p-2 text-[#141414]/40 hover:text-[#141414] transition-colors"><Filter size={20} /></button>
          </div>
        </header>

        {/* Tab Viewport */}
        <div className="flex-1 overflow-y-auto p-10 bg-[#FDFCFB]">
          <AnimatePresence mode="wait">
            {activeTab === 'icsr' && (
              <motion.div
                key="icsr"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-10"
              >
                <div className="grid grid-cols-3 gap-6">
                  <div className="col-span-2 bg-white p-8 rounded-3xl border border-[#141414]/10 shadow-sm h-[600px]">
                      <div className="flex justify-between items-end mb-8">
                        <div>
                          <h2 className="text-2xl font-bold tracking-tight">{activeTab === 'icsr' ? t.tabs.overview : t.tabs.disprop}</h2>
                          <p className="text-sm text-[#141414]/60">{t.dashboard.breakdown}</p>
                        </div>
                        {Number(denominator) > 0 && (
                          <div className="flex items-center gap-2 px-3 py-1 bg-[#1E7FB8]/10 rounded-full border border-[#1E7FB8]/20">
                            <span className="text-[10px] font-bold text-[#1E7FB8] uppercase tracking-wider">{t.dashboard.ratePer}</span>
                          </div>
                        )}
                      </div>
                    
                    <ResponsiveContainer width="100%" height="90%">
                      <ComposedChart data={overviewChartData} layout="vertical" margin={{ left: 100, right: 60, top: 20 }}>
                        <XAxis type="number" />
                        <XAxis 
                          xAxisId="secondary" 
                          type="number"
                          orientation="top" 
                          axisLine={false}
                          tickLine={false}
                          tick={{ fontSize: 10, fill: '#5BC0BE' }}
                          hide={!overviewChartData.some(d => d.frequency !== undefined)}
                          label={{ value: 'Rate / 1M', position: 'insideTopRight', offset: 0, fontSize: 10, fill: '#5BC0BE' }}
                        />
                        <YAxis 
                          type="category" 
                          dataKey="dimension" 
                          width={100} 
                          axisLine={false} 
                          tickLine={false}
                          tick={{ fontSize: 11, fontWeight: 500 }}
                        />
                        <Tooltip 
                          cursor={{ fill: 'rgba(20,20,20,0.02)' }}
                          contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 30px rgba(0,0,0,0.1)' }}
                        />
                        <Legend verticalAlign="top" align="right" />
                        <Bar 
                          dataKey="nonSerious" 
                          stackId="a"
                          fill="#92D050" 
                          name="Non serious"
                          barSize={20}
                        />
                        <Bar 
                          dataKey="serious" 
                          stackId="a"
                          fill="red" 
                          name="Serious"
                          barSize={20}
                          label={(props: any) => {
                            const { x, y, width, height, payload } = props;
                            if (!payload || payload.total === undefined) return null;
                            return (
                              <text 
                                x={x + width + 5} 
                                y={y + height / 2} 
                                dy=".35em"
                                fill="#1E7FB8" 
                                fontSize={10} 
                                fontWeight="bold"
                              >
                                {payload.total}
                              </text>
                            );
                          }}
                        />
                        {overviewChartData.some(d => d.frequency !== undefined) && (
                          <Line 
                            xAxisId="secondary"
                            dataKey="frequency" 
                            stroke="#5BC0BE" 
                            strokeWidth={2} 
                            dot={{ fill: '#5BC0BE', r: 4 }} 
                            name={`Freq / ${denomType}`}
                          />
                        )}
                      </ComposedChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="space-y-6">
                    <div className="bg-[#1E7FB8] text-white p-8 rounded-3xl shadow-xl shadow-[#1E7FB8]/10">
                      <h3 className="text-sm font-bold opacity-60 uppercase tracking-widest mb-4">Quick Stats</h3>
                      <div className="space-y-6">
                        <div>
                          <p className="text-3xl font-bold tracking-tighter">
                            {overviewData.reduce((acc, curr) => acc + curr.total, 0).toLocaleString()}
                          </p>
                          <p className="text-[11px] font-medium opacity-50 uppercase tracking-wider">Total Reports</p>
                        </div>
                        <div>
                          <p className="text-3xl font-bold tracking-tighter">
                            {overviewData.reduce((acc, curr) => acc + curr.serious, 0).toLocaleString()}
                          </p>
                          <p className="text-[11px] font-medium opacity-50 uppercase tracking-wider">Serious Reports</p>
                        </div>
                        <div>
                          <p className="text-3xl font-bold tracking-tighter">
                            {overviewData.length}
                          </p>
                          <p className="text-[11px] font-medium opacity-50 uppercase tracking-wider">Distinct Dimensions</p>
                        </div>
                        {overviewChartData[0] && (
                          <div className="pt-6 border-t border-white/10">
                            <p className="text-lg font-bold truncate">{overviewChartData[0].dimension}</p>
                            <p className="text-[11px] font-medium opacity-50 uppercase tracking-wider">Top Reported Path</p>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>

                {/* Map Section - Conditional toggle based on "Countries" sheet */}
                {selectedOverviewSheet.toLowerCase().includes('countr') && (
                  <motion.div
                    initial={{ opacity: 0, scale: 0.98 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="w-full"
                  >
                    <WorldMap data={overviewData} lang={language} />
                  </motion.div>
                )}
              </motion.div>
            )}

            {activeTab === 'disprop' && (
              <motion.div
                key="disprop"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-8"
              >
                <div className="bg-white p-8 rounded-3xl border border-[#141414]/10 shadow-sm h-[800px]">
                  <div className="mb-8">
                    <h2 className="text-2xl font-bold tracking-tight">IC & IC025 Profile</h2>
                    <p className="text-sm text-[#141414]/60">Disproportionality analysis across Preferred Terms.</p>
                  </div>
                  
                  <ResponsiveContainer width="100%" height="90%">
                    <BarChart data={dispropChartData} layout="vertical" margin={{ left: 150, right: 30 }} barGap={-12}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="rgba(20,20,20,0.05)" />
                      <XAxis type="number" />
                      <YAxis 
                        dataKey="reaction" 
                        type="category" 
                        width={150} 
                        axisLine={false} 
                        tickLine={false}
                        tick={{ fontSize: 10, fontWeight: 500 }}
                      />
                      <Tooltip 
                         contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 30px rgba(0,0,0,0.1)' }}
                      />
                      <Legend verticalAlign="top" align="right" />
                      <Bar 
                        dataKey="ic025" 
                        radius={[0, 0, 0, 0]} 
                        name="IC025"
                        barSize={12}
                      >
                        {dispropChartData.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={entry.isSignal ? '#F0A082' : '#A7E6F0'} />
                        ))}
                      </Bar>
                      <Bar 
                        dataKey="ic" 
                        fill="#1E7FB8"
                        radius={[0, 0, 0, 0]} 
                        name="IC"
                        barSize={4}
                        opacity={0.8}
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="bg-white rounded-3xl border border-[#141414]/10 shadow-sm overflow-hidden">
                  <div className="px-8 py-6 border-b border-[#141414]/10 bg-gray-50/50 flex justify-between items-center">
                    <h3 className="font-bold tracking-tight">{t.dashboard.signalTable}</h3>
                    <div className="flex items-center gap-2 text-xs text-[#1E7FB8] font-bold uppercase tracking-wider">
                      <span className="w-2 h-2 rounded-full bg-[#1E7FB8]" />
                      {dispropData.filter(d => d.isSignal).length} {t.dashboard.activeSignals}
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-gray-50/50 text-[11px] font-bold uppercase tracking-wider text-[#141414]/50">
                        <tr>
                          <th className="px-8 py-4">Reaction (PT)</th>
                          <th className="px-8 py-4">IC</th>
                          <th className="px-8 py-4">IC025</th>
                          <th className="px-8 py-4">N Observed</th>
                          <th className="px-8 py-4">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#141414]/5">
                        {dispropData.filter(d => d.isSignal).map((row, i) => (
                          <tr key={i} className="hover:bg-gray-50 transition-colors">
                            <td className="px-8 py-4 font-bold">{row.reaction}</td>
                            <td className="px-8 py-4 font-mono">{row.ic.toFixed(2)}</td>
                            <td className="px-8 py-4 font-mono">{row.ic025.toFixed(2)}</td>
                            <td className="px-8 py-4">{row.nObserved}</td>
                            <td className="px-8 py-4">
                              <span className="px-2 py-1 bg-[#1E7FB8]/10 text-[#1E7FB8] text-[10px] font-bold rounded-full border border-[#1E7FB8]/20 uppercase">
                                {t.dashboard.detected}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </motion.div>
            )}

            {activeTab === 'forest' && (
              <motion.div
                key="forest"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-8"
              >
                <div className="bg-white p-8 rounded-3xl border border-[#141414]/10 shadow-sm min-h-[600px]">
                  <div className="flex justify-between items-start mb-8">
                    <div>
                      <h2 className="text-2xl font-bold tracking-tight">
                        {selectedSubgroupSheet || "Subgroup Analysis"}
                      </h2>
                      <p className="text-sm text-[#141414]/60 max-w-2xl">{subgroupMeta || "Confidence intervals for disproportionality indices."}</p>
                    </div>
                  </div>
                  
                  {subgroupData.length > 0 ? (
                    <ResponsiveContainer width="100%" height={Math.max(500, subgroupData.length * 40)}>
                      <ComposedChart data={subgroupData} layout="vertical" margin={{ left: 150, right: 50 }}>
                        <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="rgba(20,20,20,0.05)" />
                        <XAxis type="number" domain={['auto', 'auto']} />
                        <YAxis 
                          dataKey="group" 
                          type="category" 
                          width={150} 
                          axisLine={false} 
                          tickLine={false}
                          tick={{ fontSize: 10, fontWeight: 500 }}
                        />
                        <Tooltip 
                           contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 30px rgba(0,0,0,0.1)' }}
                        />
                        <Scatter dataKey="ic" fill="#141414">
                          {subgroupData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.isSignal ? '#1E7FB8' : '#141414'} />
                          ))}
                          <ErrorBar dataKey="ic" direction="x" strokeWidth={2} stroke="#141414" opacity={0.3} xError={['ic0005', 'ic9995']} />
                        </Scatter>
                        <Legend />
                        {/* Vertical line at 0 manually simulated if needed, or rely on axis */}
                      </ComposedChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="flex flex-col items-center justify-center h-[400px] text-[#141414]/20">
                      <Trees size={64} className="mb-4" />
                      <p className="text-lg font-bold italic serif">No subgroup data loaded</p>
                    </div>
                  )}
                </div>
              </motion.div>
            )}

            {activeTab === 'synonyms' && (
              <motion.div
                key="synonyms"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-8"
              >
                <div className="max-w-5xl">
                  <h2 className="text-3xl font-bold tracking-tight mb-2">{t.ai.title}</h2>
                  <p className="text-[#141414]/60 mb-8 italic serif text-lg underline underline-offset-8 decoration-[#1E7FB8]/30">
                    {t.ai.subtitle}
                  </p>
                  
                  {dispropData.length > 0 ? (
                    <div className="space-y-6">
                      <div className="flex items-center justify-between bg-[#141414]/5 p-6 rounded-3xl border border-[#141414]/10">
                        <div className="space-y-1">
                          <p className="text-sm font-bold uppercase tracking-wider opacity-60">Analysis Scope</p>
                          <p className="text-sm">Processing top 30 Preferred Terms for potential duplicates and synonyms.</p>
                        </div>
                        <button
                          onClick={fetchSynonyms}
                          disabled={isAiLoading}
                          className="flex items-center gap-2 px-6 py-3 bg-[#141414] text-white rounded-2xl font-bold shadow-xl shadow-[#141414]/20 hover:bg-[#1E7FB8] transition-all disabled:opacity-50"
                        >
                          {isAiLoading ? <Loader2 className="animate-spin" /> : <Languages size={18} />}
                          {isAiLoading ? t.ai.loading : t.ai.button}
                        </button>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {synonymGroups.map((group, i) => (GroupCard(group, i, language)))}
                      </div>
                    </div>
                  ) : (
                    <div className="p-12 border-2 border-dashed border-[#141414]/10 rounded-3xl text-center bg-gray-50/50">
                      <Languages size={48} className="mx-auto mb-4 opacity-20" />
                      <p className="text-lg font-bold opacity-40 italic">Upload Disproportionality file to enable AI synonym grouping.</p>
                    </div>
                  )}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </main>
    </div>
  );
}
