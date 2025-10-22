import React, { useState, useCallback, useRef, useEffect } from 'react';
import { GoogleGenAI, Type } from '@google/genai';

// Extend the Window interface to include globally loaded libraries for TypeScript
interface CustomWindow extends Window {
  XLSX: any;
  Chart: any;
  jspdf: {
    jsPDF: new (options?: any) => any;
  };
  html2canvas: (element: HTMLElement, options?: any) => Promise<HTMLCanvasElement>;
}

declare let window: CustomWindow;

interface IAnalysisResult {
  eqp: string;
  count: number;
}

interface IChartData {
  labels: string[];
  datasets: any[];
}

interface IAIDiagnosis {
    equipamento: string;
    padraoAtuacao: string;
    interpretacaoTecnica: string;
    acaoRecomendada: string;
}

const App: React.FC = () => {
  const [analysisResult, setAnalysisResult] = useState<IAnalysisResult[]>([]);
  const [chartData, setChartData] = useState<IChartData | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  
  // AI State
  const [isDiagnosing, setIsDiagnosing] = useState<boolean>(false);
  const [aiDiagnosis, setAiDiagnosis] = useState<IAIDiagnosis[] | null>(null);
  const [aiError, setAiError] = useState<string | null>(null);


  const chartRef = useRef<HTMLCanvasElement | null>(null);
  const chartInstanceRef = useRef<any | null>(null);
  const resultsContainerRef = useRef<HTMLDivElement | null>(null);
  const chartContainerRef = useRef<HTMLDivElement | null>(null);
  const diagnosisContainerRef = useRef<HTMLDivElement | null>(null);

  const processFile = useCallback((file: File) => {
    setIsLoading(true);
    setError(null);
    setAiError(null);
    setAnalysisResult([]);
    setChartData(null);
    setAiDiagnosis(null);
    setFileName(file.name);

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = window.XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json: any[] = window.XLSX.utils.sheet_to_json(worksheet);

        if (json.length === 0) {
          throw new Error("A planilha está vazia ou em um formato incorreto.");
        }

        const firstRow = json[0];
        if (!('EQP' in firstRow) || !('Data' in firstRow) || !('Hora' in firstRow)) {
          throw new Error("A planilha deve conter as colunas 'EQP', 'Data' e 'Hora'.");
        }
        
        const actuationCounts: { [key: string]: number } = {};
        
        for (const row of json) {
          const eqp = row['EQP'];
          if(eqp) {
            actuationCounts[eqp] = (actuationCounts[eqp] || 0) + 1;
          }
        }
        
        const sortedResults = Object.entries(actuationCounts)
          .map(([eqp, count]) => ({ eqp, count }))
          .sort((a, b) => b.count - a.count)
          .slice(0, 10);
        
        if (sortedResults.length === 0) {
           throw new Error("Nenhum dado de 'EQP' válido foi encontrado para análise.");
        }

        setAnalysisResult(sortedResults);

        // --- Chart Data Processing ---
        const top3Eqps = sortedResults.slice(0, 3).map(r => r.eqp);
        const hourlyCounts: { [eqp: string]: number[] } = {};
        top3Eqps.forEach(eqp => {
          hourlyCounts[eqp] = Array(24).fill(0);
        });

        for (const row of json) {
            const eqp = row['EQP'];
            if (top3Eqps.includes(eqp)) {
                const time = row['Hora'];
                let hour: number | null = null;
                if (typeof time === 'string') {
                    const hourPart = time.split(':')[0];
                    const parsedHour = parseInt(hourPart, 10);
                    if (!isNaN(parsedHour)) {
                        hour = parsedHour;
                    }
                } else if (typeof time === 'number' && time < 1) {
                    hour = Math.floor(time * 24);
                }

                if (hour !== null && hour >= 0 && hour < 24) {
                    hourlyCounts[eqp][hour]++;
                }
            }
        }
        
        const labels = Array.from({ length: 24 }, (_, i) => `${String(i).padStart(2, '0')}:00`);
        const colors = ['rgba(0, 123, 255, 1)', 'rgba(40, 167, 69, 1)', 'rgba(255, 193, 7, 1)'];
        
        const datasets = top3Eqps.map((eqp, index) => ({
            label: `EQP: ${eqp}`,
            data: hourlyCounts[eqp],
            borderColor: colors[index],
            backgroundColor: colors[index].replace('1)', '0.2)'),
            fill: false,
            tension: 0.1
        }));

        setChartData({ labels, datasets });

      } catch (err) {
        if (err instanceof Error) {
            setError(`Erro ao processar o arquivo: ${err.message}`);
        } else {
            setError("Ocorreu um erro desconhecido ao processar o arquivo.");
        }
      } finally {
        setIsLoading(false);
      }
    };

    reader.onerror = () => {
        setError("Não foi possível ler o arquivo selecionado.");
        setIsLoading(false);
    }

    reader.readAsArrayBuffer(file);
  }, []);
  
  useEffect(() => {
    if (chartRef.current && chartData) {
      if (chartInstanceRef.current) {
        chartInstanceRef.current.destroy();
      }

      const ctx = chartRef.current.getContext('2d');
      if (ctx) {
        chartInstanceRef.current = new window.Chart(ctx, {
          type: 'line',
          data: chartData,
          options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { position: 'top' },
              tooltip: { mode: 'index', intersect: false },
            },
            scales: {
              y: {
                beginAtZero: true,
                title: { display: true, text: 'Quantidade de Atuações' }
              },
              x: {
                title: { display: true, text: 'Hora do Dia' }
              }
            }
          },
        });
      }
    }
    return () => {
        if (chartInstanceRef.current) {
            chartInstanceRef.current.destroy();
        }
    }
  }, [chartData]);

  const handleGenerateDiagnosis = async () => {
    if (!analysisResult.length || !chartData) {
        setAiError("Dados insuficientes para gerar diagnóstico.");
        return;
    }

    setIsDiagnosing(true);
    setAiError(null);
    setAiDiagnosis(null);

    try {
        const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
        
        const top3Summary = analysisResult.slice(0, 3).map((r, i) => 
            `- ${r.eqp}: ${r.count} atuações. Padrão horário (array 24h): ${JSON.stringify(chartData.datasets[i]?.data)}`
        ).join('\n');

        const prompt = `
            Você é um especialista em manutenção preditiva para equipamentos pesados de movimentação de granéis, como transportadores de correia, carregadores de navio, recuperadoras e empilhadeiras-recuperadoras.
            Analise os seguintes dados de atuação dos 3 principais equipamentos e forneça um diagnóstico técnico e acionável.

            Dados de Atuação:
            ${top3Summary}

            Para cada um dos 3 equipamentos, retorne um objeto JSON com as seguintes chaves:
            - "equipamento": Nome do equipamento.
            - "padraoAtuacao": Descrição curta do padrão de atuações (ex: "Picos noturnos", "Operação constante", "Atividade concentrada na madrugada").
            - "interpretacaoTecnica": Diagnóstico técnico conciso, relacionando o padrão a possíveis causas em equipamentos como estes (ex: "Desgaste de componentes mecânicos", "Falha em sensores", "Sobrecarga operacional em horários específicos").
            - "acaoRecomendada": Sugestão de ação clara e direta para a equipe de manutenção (ex: "Inspecionar sistema de tração", "Verificar alinhamento da correia transportadora", "Analisar alarmes do período de pico").

            O resultado final deve ser um array de objetos JSON.
        `;
        
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: {
                responseMimeType: 'application/json',
                responseSchema: {
                    type: Type.ARRAY,
                    items: {
                        type: Type.OBJECT,
                        properties: {
                            equipamento: { type: Type.STRING },
                            padraoAtuacao: { type: Type.STRING, description: "Descrição curta do padrão (ex: picos noturnos, constante)." },
                            interpretacaoTecnica: { type: Type.STRING, description: "Diagnóstico técnico curto (ex: desgaste esperado, falha iminente)." },
                            acaoRecomendada: { type: Type.STRING, description: "Ação imediata sugerida (ex: inspecionar, agendar manutenção)." }
                        },
                        required: ["equipamento", "padraoAtuacao", "interpretacaoTecnica", "acaoRecomendada"]
                    }
                }
            }
        });

        const diagnosisData = JSON.parse(response.text);
        setAiDiagnosis(diagnosisData);

    } catch (err) {
        console.error("AI Diagnosis Error:", err);
        setAiError("Não foi possível gerar o diagnóstico. Tente novamente.");
    } finally {
        setIsDiagnosing(false);
    }
  };


  const handleExportPDF = async () => {
    if (!resultsContainerRef.current || !chartContainerRef.current) {
      setError("Elementos do relatório não encontrados para exportação.");
      return;
    }
    
    setIsExporting(true);
    setError(null);
    
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
            orientation: 'p',
            unit: 'mm',
            format: 'a4',
        });
        
        const margin = 10;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - margin * 2;

        // 1. Capture and add Table
        const tableCanvas = await window.html2canvas(resultsContainerRef.current, { scale: 2 });
        const tableImgData = tableCanvas.toDataURL('image/png');
        const tableImgProps = doc.getImageProperties(tableImgData);
        const tableImgHeight = (tableImgProps.height * contentWidth) / tableImgProps.width;
        doc.addImage(tableImgData, 'PNG', margin, margin, contentWidth, tableImgHeight);

        // 2. Add new page for the chart
        doc.addPage();
        
        // 3. Capture and add Chart
        const chartCanvas = await window.html2canvas(chartContainerRef.current, { scale: 2 });
        const chartImgData = chartCanvas.toDataURL('image/png');
        const chartImgProps = doc.getImageProperties(chartImgData);
        const chartImgHeight = (chartImgProps.height * contentWidth) / chartImgProps.width;
        doc.addImage(chartImgData, 'PNG', margin, margin, contentWidth, chartImgHeight);
        
        // 4. Capture and add AI Diagnosis if it exists
        if (aiDiagnosis && diagnosisContainerRef.current) {
            doc.addPage();
            const diagnosisCanvas = await window.html2canvas(diagnosisContainerRef.current, { scale: 2 });
            const diagnosisImgData = diagnosisCanvas.toDataURL('image/png');
            const diagnosisImgProps = doc.getImageProperties(diagnosisImgData);
            const diagnosisImgHeight = (diagnosisImgProps.height * contentWidth) / diagnosisImgProps.width;
            doc.addImage(diagnosisImgData, 'PNG', margin, margin, contentWidth, diagnosisImgHeight);
        }
        
        doc.save(`Relatorio_${fileName || 'atuacoes'}.pdf`);
    } catch (err) {
        setError("Ocorreu um erro ao gerar o PDF.");
        console.error(err);
    } finally {
        setIsExporting(false);
    }
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.xlsx')) {
        processFile(file);
      } else {
        setError("Por favor, selecione um arquivo .xlsx");
      }
    }
    event.target.value = '';
  };

  return (
    <div className="container">
      <h1 className="title">Analisador de Atuações de Equipamentos</h1>
      <p className="description">
        Importe uma planilha (.xlsx) com as colunas "EQP", "Data" e "Hora" para analisar a quantidade de atuações e visualizar tendências.
      </p>

      <div className="file-uploader">
        <label htmlFor="file-input" className="file-label">
          Selecionar Planilha (.xlsx)
        </label>
        <input
          id="file-input"
          type="file"
          className="file-input"
          accept=".xlsx"
          onChange={handleFileChange}
          aria-label="Seletor de arquivo de planilha"
        />
      </div>
      
      {isLoading && <div className="status-message loading">Processando "{fileName}"...</div>}
      {error && <div className="status-message error">{error}</div>}

      {analysisResult.length > 0 && (
        <div ref={resultsContainerRef} className="results-container" role="region" aria-labelledby="results-title">
          <h2 id="results-title" className="results-title">Top 10 Equipamentos com Mais Atuações</h2>
          <table className="results-table">
            <thead>
              <tr>
                <th scope="col">Equipamento (EQP)</th>
                <th scope="col">Quantidade de Atuações</th>
              </tr>
            </thead>
            <tbody>
              {analysisResult.map((result) => (
                <tr key={result.eqp}>
                  <td>{result.eqp}</td>
                  <td>{result.count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {chartData && (
        <div ref={chartContainerRef} className="chart-container">
            <h2 className="chart-title">Tendência de Atuações por Hora (Top 3)</h2>
            <canvas ref={chartRef} aria-label="Gráfico de linha de atuações por hora" role="img"></canvas>
        </div>
      )}

        {analysisResult.length > 0 && (
            <div className="actions-container">
                <button
                    onClick={handleGenerateDiagnosis}
                    className="action-button primary"
                    disabled={isDiagnosing}
                >
                    {isDiagnosing ? 'Analisando...' : 'Gerar Diagnóstico Preditivo (IA)'}
                </button>
                <button 
                    onClick={handleExportPDF} 
                    className="action-button success" 
                    disabled={isExporting}
                >
                    {isExporting ? 'Exportando PDF...' : 'Exportar Relatório em PDF'}
                </button>
            </div>
        )}

      {isDiagnosing && <div className="status-message loading">A IA está analisando os dados...</div>}
      {aiError && <div className="status-message error">{aiError}</div>}

      {aiDiagnosis && (
        <div ref={diagnosisContainerRef} className="diagnosis-container">
          <h2 className="results-title">Diagnóstico Preditivo por IA</h2>
          <table className="diagnosis-table">
            <thead>
              <tr>
                <th>Equipamento</th>
                <th>Padrão de Atuação</th>
                <th>Interpretação Técnica</th>
                <th>Ação Recomendada</th>
              </tr>
            </thead>
            <tbody>
              {aiDiagnosis.map((diag, index) => (
                <tr key={index}>
                  <td>{diag.equipamento}</td>
                  <td>{diag.padraoAtuacao}</td>
                  <td>{diag.interpretacaoTecnica}</td>
                  <td>{diag.acaoRecomendada}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default App;
