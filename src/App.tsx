import React, { useState, useCallback } from 'react';
import { GoogleGenAI, Type } from '@google/genai';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { UploadCloud, Activity, Database, Zap, FileText, BarChart3, PieChart, TrendingUp, AlertCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  BarChart, Bar, PieChart as RechartsPieChart, Pie, Cell, Legend
} from 'recharts';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
interface AnalysisResult {
  context: {
    classification: string;
    overview: string;
  };
  metrics: Array<{
    label: string;
    value: string;
    trend: string;
  }>;
  insights: string[];
  recommendations: string[];
  visualization: {
    trend_data: Array<{ label: string; value: number }>;
    composition_data: Array<{ label: string; value: number }>;
    comparison_data: Array<{ label: string; value: number }>;
  };
}

// --- Components ---

const Card: React.FC<{ children: React.ReactNode; className?: string }> = ({ children, className }) => (
  <div className={cn("bg-[#151619] border border-white/10 rounded-xl p-6 shadow-2xl", className)}>
    {children}
  </div>
);

const SectionHeader = ({ icon: Icon, title, subtitle }: { icon: any, title: string, subtitle?: string }) => (
  <div className="flex items-center gap-3 mb-6 border-b border-white/10 pb-4">
    <div className="p-2 bg-emerald-500/10 rounded-lg border border-emerald-500/20">
      <Icon className="w-5 h-5 text-emerald-400" />
    </div>
    <div>
      <h2 className="text-lg font-mono font-semibold text-white tracking-tight uppercase">{title}</h2>
      {subtitle && <p className="text-xs text-gray-400 font-mono mt-1">{subtitle}</p>}
    </div>
  </div>
);

const COLORS = ['#10b981', '#3b82f6', '#8b5cf6', '#f59e0b', '#ef4444', '#ec4899'];

export default function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [fileInfo, setFileInfo] = useState<{ name: string; rows: number; columns: number } | null>(null);

  const processFile = async (file: File) => {
    setError(null);
    setIsAnalyzing(true);
    setAnalysisResult(null);

    try {
      let data: any[] = [];
      let headers: string[] = [];

      if (file.name.endsWith('.csv')) {
        const text = await file.text();
        const result = Papa.parse(text, { header: true, skipEmptyLines: true });
        data = result.data;
        headers = result.meta.fields || [];
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        data = XLSX.utils.sheet_to_json(worksheet);
        if (data.length > 0) {
          headers = Object.keys(data[0] as object);
        }
      } else if (file.name.endsWith('.json')) {
        const text = await file.text();
        const parsed = JSON.parse(text);
        data = Array.isArray(parsed) ? parsed : [parsed];
        if (data.length > 0) {
          headers = Object.keys(data[0] as object);
        }
      } else {
        throw new Error("Unsupported file format. Please upload CSV, Excel, or JSON.");
      }

      if (data.length === 0) {
        throw new Error("File is empty or format is invalid.");
      }

      setFileInfo({
        name: file.name,
        rows: data.length,
        columns: headers.length,
      });

      // Prepare sample for Gemini
      const sampleSize = Math.min(data.length, 50);
      const sampleData = data.slice(0, sampleSize);

      await analyzeDataWithGemini(sampleData, headers, data.length);

    } catch (err: any) {
      console.error(err);
      setError(err.message || "Error processing file. Please try again.");
      setIsAnalyzing(false);
    }
  };

  const analyzeDataWithGemini = async (sampleData: any[], headers: string[], totalRows: number) => {
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      
      const prompt = `
Act as the "Pami Automate Universal Intelligence Engine." Your goal is to function as an advanced Data Scientist and UI/UX Strategy Expert capable of analyzing ANY dataset provided.

CONTEXT:
Here is a sample of the dataset (first ${sampleData.length} rows) and its schema:
Headers: ${headers.join(', ')}
Sample Data:
${JSON.stringify(sampleData, null, 2)}

TASK:
Analyze the uploaded file and provide a structured JSON response to populate a dynamic infographic dashboard. Your analysis must cover:

DATA CONTEXT & CLASSIFICATION:
- Identify the nature of the dataset (e.g., "Dataset ini dikesan sebagai rekod Transaksi Jualan").
- Total records: ${totalRows}
- Number of variables/columns: ${headers.length}
- Brief overview of the data's scope and timeframe.

AUTOMATED QUANTITATIVE METRICS:
- Identify the top 3-4 most critical KPIs found in the data.
- Provide the calculated values for these metrics.
- Compare these values to identify trends or anomalies.

STRATEGIC INSIGHTS & SYNTHESIS:
- Provide 3 deep insights based on correlations found.
- Give 3 "Cyber-Industrial" recommendations (Projeksi Strategik) for optimization based on the findings.

DATA FOR VISUALIZATION:
- trend_data: A time-series or sequence projection (labels and values).
- composition_data: Breakdown of categories.
- comparison_data: Ranking of the top-performing entities/categories in the data.

TONE & LANGUAGE:
All narrative text MUST be in professional English.
Use "Cyber-Industrial" terminology (e.g., 'Data Synthesis', 'Intel Automata', 'Operational Core').

CONSTRAINTS:
Format all financial values in RM (if currency is detected) or appropriate units.
Return ONLY a valid JSON object matching the requested schema.
`;

      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: prompt,
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              context: {
                type: Type.OBJECT,
                properties: {
                  classification: { type: Type.STRING },
                  overview: { type: Type.STRING }
                },
                required: ["classification", "overview"]
              },
              metrics: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    label: { type: Type.STRING },
                    value: { type: Type.STRING },
                    trend: { type: Type.STRING }
                  },
                  required: ["label", "value", "trend"]
                }
              },
              insights: {
                type: Type.ARRAY,
                items: { type: Type.STRING }
              },
              recommendations: {
                type: Type.ARRAY,
                items: { type: Type.STRING }
              },
              visualization: {
                type: Type.OBJECT,
                properties: {
                  trend_data: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        label: { type: Type.STRING },
                        value: { type: Type.NUMBER }
                      },
                      required: ["label", "value"]
                    }
                  },
                  composition_data: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        label: { type: Type.STRING },
                        value: { type: Type.NUMBER }
                      },
                      required: ["label", "value"]
                    }
                  },
                  comparison_data: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        label: { type: Type.STRING },
                        value: { type: Type.NUMBER }
                      },
                      required: ["label", "value"]
                    }
                  }
                },
                required: ["trend_data", "composition_data", "comparison_data"]
              }
            },
            required: ["context", "metrics", "insights", "recommendations", "visualization"]
          }
        }
      });

      const jsonStr = response.text;
      if (jsonStr) {
        const result = JSON.parse(jsonStr) as AnalysisResult;
        setAnalysisResult(result);
      } else {
        throw new Error("Failed to get response from Pami Automate.");
      }
    } catch (err: any) {
      console.error(err);
      setError("Error during data synthesis: " + (err.message || "Please try again."));
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  }, []);

  const handleFileInput = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  }, []);

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-gray-200 font-sans selection:bg-emerald-500/30">
      {/* Header */}
      <header className="border-b border-white/5 bg-[#0a0a0a]/80 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded bg-emerald-500 flex items-center justify-center shadow-[0_0_15px_rgba(16,185,129,0.4)]">
              <Zap className="w-5 h-5 text-[#0a0a0a]" />
            </div>
            <div>
              <h1 className="font-mono font-bold text-white tracking-tight">PAMI AUTOMATE</h1>
              <p className="text-[10px] font-mono text-emerald-500 tracking-widest uppercase">Universal Intelligence Engine</p>
            </div>
          </div>
          <div className="flex items-center gap-4 text-sm font-mono text-gray-500">
            <span className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
              System Active
            </span>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-8">
        {!analysisResult && !isAnalyzing && (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="max-w-2xl mx-auto mt-12"
          >
            <div className="text-center mb-8">
              <h2 className="text-3xl font-mono font-bold text-white mb-4">Initialize Operational Core</h2>
              <p className="text-gray-400">Upload your dataset to begin Intel Automata and Data Synthesis.</p>
            </div>

            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={cn(
                "relative group cursor-pointer rounded-2xl border-2 border-dashed transition-all duration-300 p-12 text-center",
                isDragging 
                  ? "border-emerald-500 bg-emerald-500/5" 
                  : "border-white/10 bg-[#151619] hover:border-emerald-500/50 hover:bg-[#1a1b1e]"
              )}
            >
              <input
                type="file"
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                onChange={handleFileInput}
                accept=".csv, .xlsx, .xls, .json"
              />
              <div className="flex flex-col items-center gap-4">
                <div className={cn(
                  "p-4 rounded-xl transition-colors duration-300",
                  isDragging ? "bg-emerald-500/20 text-emerald-400" : "bg-white/5 text-gray-400 group-hover:text-emerald-400"
                )}>
                  <UploadCloud className="w-8 h-8" />
                </div>
                <div>
                  <p className="text-lg font-medium text-white mb-1">Drag & Drop File Here</p>
                  <p className="text-sm text-gray-500">Supported formats: CSV, Excel (.xlsx), JSON</p>
                </div>
              </div>
            </div>

            {error && (
              <div className="mt-6 p-4 rounded-lg bg-red-500/10 border border-red-500/20 flex items-start gap-3 text-red-400">
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <p className="text-sm">{error}</p>
              </div>
            )}
          </motion.div>
        )}

        {isAnalyzing && (
          <div className="flex flex-col items-center justify-center py-32">
            <div className="relative w-24 h-24 mb-8">
              <div className="absolute inset-0 rounded-full border-t-2 border-emerald-500 animate-spin"></div>
              <div className="absolute inset-2 rounded-full border-r-2 border-blue-500 animate-spin-reverse"></div>
              <div className="absolute inset-4 rounded-full border-b-2 border-purple-500 animate-spin"></div>
              <Activity className="absolute inset-0 m-auto w-8 h-8 text-emerald-400 animate-pulse" />
            </div>
            <h3 className="text-xl font-mono font-bold text-white mb-2">Processing Operational Core...</h3>
            <p className="text-gray-400 font-mono text-sm animate-pulse">Data Synthesis in Progress</p>
            
            {fileInfo && (
              <div className="mt-8 flex gap-6 text-sm font-mono text-gray-500">
                <span className="flex items-center gap-2"><FileText className="w-4 h-4" /> {fileInfo.name}</span>
                <span className="flex items-center gap-2"><Database className="w-4 h-4" /> {fileInfo.rows} Records</span>
                <span className="flex items-center gap-2"><BarChart3 className="w-4 h-4" /> {fileInfo.columns} Variables</span>
              </div>
            )}
          </div>
        )}

        {analysisResult && !isAnalyzing && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="space-y-6"
          >
            {/* Context & Top Metrics */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <Card className="lg:col-span-1 flex flex-col justify-center">
                <SectionHeader icon={Database} title="Data Classification" subtitle="Context Synthesis" />
                <div className="space-y-4">
                  <div>
                    <p className="text-xs font-mono text-gray-500 uppercase mb-1">Dataset Identity</p>
                    <p className="text-lg font-medium text-emerald-400">{analysisResult.context.classification}</p>
                  </div>
                  <div>
                    <p className="text-xs font-mono text-gray-500 uppercase mb-1">Operational Scope</p>
                    <p className="text-sm text-gray-300 leading-relaxed">{analysisResult.context.overview}</p>
                  </div>
                  {fileInfo && (
                     <div className="pt-4 border-t border-white/5 flex justify-between text-xs font-mono text-gray-500">
                       <span>{fileInfo.rows} RECORDS</span>
                       <span>{fileInfo.columns} VARIABLES</span>
                     </div>
                  )}
                </div>
              </Card>

              <div className="lg:col-span-2 grid grid-cols-1 sm:grid-cols-3 gap-4">
                {analysisResult.metrics.map((metric, idx) => (
                  <Card key={idx} className="flex flex-col justify-between">
                    <p className="text-xs font-mono text-gray-500 uppercase mb-2">{metric.label}</p>
                    <p className="text-3xl font-mono text-white mb-4">{metric.value}</p>
                    <div className="flex items-center gap-2 text-sm">
                      <TrendingUp className="w-4 h-4 text-emerald-500" />
                      <span className="text-gray-400">{metric.trend}</span>
                    </div>
                  </Card>
                ))}
              </div>
            </div>

            {/* Visualizations */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <Card>
                <SectionHeader icon={TrendingUp} title="Trend Projection" subtitle="Time-Series Analysis" />
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={analysisResult.visualization.trend_data}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#333" vertical={false} />
                      <XAxis dataKey="label" stroke="#666" fontSize={12} tickLine={false} axisLine={false} />
                      <YAxis stroke="#666" fontSize={12} tickLine={false} axisLine={false} tickFormatter={(val) => `${val}`} />
                      <RechartsTooltip 
                        contentStyle={{ backgroundColor: '#151619', borderColor: '#333', borderRadius: '8px' }}
                        itemStyle={{ color: '#10b981' }}
                      />
                      <Line type="monotone" dataKey="value" stroke="#10b981" strokeWidth={3} dot={{ r: 4, fill: '#151619', strokeWidth: 2 }} activeDot={{ r: 6 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </Card>

              <Card>
                <SectionHeader icon={BarChart3} title="Entity Comparison" subtitle="Top Performance Ranking" />
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={analysisResult.visualization.comparison_data} layout="vertical" margin={{ left: 20 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#333" horizontal={false} />
                      <XAxis type="number" stroke="#666" fontSize={12} tickLine={false} axisLine={false} />
                      <YAxis dataKey="label" type="category" stroke="#666" fontSize={12} tickLine={false} axisLine={false} width={100} />
                      <RechartsTooltip 
                        contentStyle={{ backgroundColor: '#151619', borderColor: '#333', borderRadius: '8px' }}
                        cursor={{ fill: '#ffffff05' }}
                      />
                      <Bar dataKey="value" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={24} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </Card>

              <Card className="lg:col-span-2">
                <SectionHeader icon={PieChart} title="Data Composition" subtitle="Category Distribution" />
                <div className="h-80">
                  <ResponsiveContainer width="100%" height="100%">
                    <RechartsPieChart>
                      <Pie
                        data={analysisResult.visualization.composition_data}
                        cx="50%"
                        cy="50%"
                        innerRadius={80}
                        outerRadius={120}
                        paddingAngle={5}
                        dataKey="value"
                        nameKey="label"
                      >
                        {analysisResult.visualization.composition_data.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Pie>
                      <RechartsTooltip 
                        contentStyle={{ backgroundColor: '#151619', borderColor: '#333', borderRadius: '8px' }}
                      />
                      <Legend verticalAlign="bottom" height={36} iconType="circle" />
                    </RechartsPieChart>
                  </ResponsiveContainer>
                </div>
              </Card>
            </div>

            {/* Insights & Recommendations */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <Card>
                <SectionHeader icon={Zap} title="Intel Synthesis" subtitle="Correlations & Key Findings" />
                <div className="space-y-4">
                  {analysisResult.insights.map((insight, idx) => (
                    <div key={idx} className="flex gap-4 items-start p-4 rounded-lg bg-white/5 border border-white/5">
                      <div className="w-6 h-6 rounded-full bg-blue-500/20 flex items-center justify-center shrink-0 mt-0.5">
                        <span className="text-xs font-mono text-blue-400">{idx + 1}</span>
                      </div>
                      <p className="text-sm text-gray-300 leading-relaxed">{insight}</p>
                    </div>
                  ))}
                </div>
              </Card>

              <Card>
                <SectionHeader icon={Activity} title="Strategic Projection" subtitle="Optimization Recommendations" />
                <div className="space-y-4">
                  {analysisResult.recommendations.map((rec, idx) => (
                    <div key={idx} className="flex gap-4 items-start p-4 rounded-lg bg-emerald-500/5 border border-emerald-500/10">
                      <div className="w-6 h-6 rounded-full bg-emerald-500/20 flex items-center justify-center shrink-0 mt-0.5">
                        <span className="text-xs font-mono text-emerald-400">{idx + 1}</span>
                      </div>
                      <p className="text-sm text-gray-300 leading-relaxed">{rec}</p>
                    </div>
                  ))}
                </div>
              </Card>
            </div>

            <div className="flex justify-center pt-8 pb-12">
              <button
                onClick={() => {
                  setAnalysisResult(null);
                  setFileInfo(null);
                }}
                className="px-6 py-3 rounded-full bg-white/5 hover:bg-white/10 border border-white/10 text-sm font-mono transition-colors"
              >
                Upload New Dataset
              </button>
            </div>
          </motion.div>
        )}
      </main>
    </div>
  );
}
