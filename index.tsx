import React, { useState, useCallback, useRef, useEffect } from 'react';
import { createRoot } from 'react-dom/client';

// Make external libraries available, as they're loaded from script tags
declare const XLSX: any;
declare const Chart: any;
declare const jspdf: any;
declare const html2canvas: any;

interface IAnalysisResult {
  eqp: string;
  count: number;
}

interface IChartData {
    labels: string[];
    datasets: {label: string; data: number[]}[];
}

const App = () => {
  const [analysisResult, setAnalysisResult] = useState<IAnalysisResult[]>([]);
  const [chartData, setChartData] = useState<IChartData | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);

  const chartRef = useRef<HTMLCanvasElement | null>(null);
  const chartInstanceRef = useRef<any | null>(null);


  const processFile = useCallback((file: File) => {
    setIsLoading(true);
    setError(null);
    setAnalysisResult([]);
    setChartData(null);
    setFileName(file.name);

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json: any[] = XLSX.utils.sheet_to_json(worksheet);

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
        if (top3Eqps.length > 0) {
            const hourlyCounts: { [key: string]: number[] } = {};
            top3Eqps.forEach(eqp => {
                hourlyCounts[eqp] = Array(24).fill(0);
            });

            for (const row of json) {
                if (top3Eqps.includes(row['EQP'])) {
                    const timeValue = row['Hora'];
                    let hour: number | null = null;
                    
                    if (typeof timeValue === 'number' && timeValue >= 0 && timeValue < 1) {
                        hour = Math.floor(timeValue * 24);
                    } else if (typeof timeValue === 'string') {
                        const hourPart = parseInt(timeValue.split(':')[0], 10);
                        if (!isNaN(hourPart) && hourPart >= 0 && hourPart < 24) {
                            hour = hourPart;
                        }
                    } else if (timeValue instanceof Date) {
                         hour = timeValue.getHours();
                    }

                    if (hour !== null) {
                        hourlyCounts[row['EQP']][hour]++;
                    }
                }
            }
            
            const chartLabels = Array.from({ length: 24 }, (_, i) => `${i}h`);
            const chartDatasets = top3Eqps.map((eqp, index) => ({
                label: eqp,
                data: hourlyCounts[eqp],
                fill: true,
                tension: 0.2
            }));

            setChartData({ labels: chartLabels, datasets: chartDatasets });
        }


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
        chartInstanceRef.current = new Chart(ctx, {
          type: 'line',
          data: chartData,
          options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: {
                position: 'top',
              },
              title: {
                display: false,
              }
            },
            scales: {
              x: {
                title: {
                  display: true,
                  text: 'Hora do Dia'
                }
              },
              y: {
                title: {
                  display: true,
                  text: 'Nº de Atuações'
                },
                beginAtZero: true
              }
            }
          }
        });
      }
    }
    return () => {
        if (chartInstanceRef.current) {
            chartInstanceRef.current.destroy();
            chartInstanceRef.current = null;
        }
    }
  }, [chartData]);


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

  const handleExportPDF = async () => {
    const tableContainer = document.getElementById('results-container');
    const chartContainer = document.getElementById('chart-container');

    if (!tableContainer) {
        setError("Não foi possível encontrar a tabela para exportar.");
        return;
    }

    setIsExporting(true);
    try {
        const { jsPDF } = jspdf;
        const pdf = new jsPDF({
            orientation: 'p',
            unit: 'mm',
            format: 'a4'
        });

        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfMargin = 10;
        const contentWidth = pdfWidth - (pdfMargin * 2);

        // --- Page 1: Table ---
        const tableCanvas = await html2canvas(tableContainer, {
            scale: 2,
            useCORS: true,
        });
        const tableImgData = tableCanvas.toDataURL('image/png');
        const tableImgProps = pdf.getImageProperties(tableImgData);
        const tableRatio = contentWidth / tableImgProps.width;
        const tableImgHeight = tableImgProps.height * tableRatio;
        
        pdf.addImage(tableImgData, 'PNG', pdfMargin, pdfMargin, contentWidth, tableImgHeight);

        // --- Page 2: Chart (if it exists) ---
        if (chartContainer) {
            pdf.addPage();
            // Wait for chart animation to complete
            await new Promise(resolve => setTimeout(resolve, 500));
            
            const chartCanvas = await html2canvas(chartContainer, {
                scale: 2,
                useCORS: true,
            });
            const chartImgData = chartCanvas.toDataURL('image/png');
            const chartImgProps = pdf.getImageProperties(chartImgData);
            const chartRatio = contentWidth / chartImgProps.width;
            const chartImgHeight = chartImgProps.height * chartRatio;

            pdf.addImage(chartImgData, 'PNG', pdfMargin, pdfMargin, contentWidth, chartImgHeight);
        }

        pdf.save(`relatorio-atuacoes-${fileName}.pdf`);

    } catch (err) {
        console.error(err);
        setError("Ocorreu um erro ao gerar o PDF.");
    } finally {
        setIsExporting(false);
    }
  };


  return (
    <div className="container">
      <h1 className="title">Analisador de Atuações de Equipamentos</h1>
      <p className="description">
        Importe uma planilha (.xlsx) com as colunas "EQP", "Data" e "Hora" para analisar a quantidade de atuações.
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

      <div id="report-content">
        {analysisResult.length > 0 && (
          <div id="results-container" className="results-container" role="region" aria-labelledby="results-title">
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
          <div id="chart-container" className="chart-container" role="region" aria-labelledby="chart-title">
              <h2 id="chart-title" className="chart-title">Gráfico de Tendência por Hora (Top 3)</h2>
              <canvas ref={chartRef}></canvas>
          </div>
        )}
      </div>

       {analysisResult.length > 0 && (
        <div className="export-container">
          <button onClick={handleExportPDF} className="export-button" disabled={isExporting}>
            {isExporting ? 'Exportando PDF...' : 'Exportar Relatório em PDF'}
          </button>
        </div>
      )}

    </div>
  );
};

const container = document.getElementById('root');
const root = createRoot(container!);
root.render(<App />);