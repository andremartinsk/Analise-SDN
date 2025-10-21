import React, { useState, useCallback, useRef, useEffect } from 'react';

// Make external libraries available in the scope, as they're loaded from script tags
declare const XLSX: any;
declare const Chart: any;

interface IAnalysisResult {
  eqp: string;
  count: number;
}

interface IChartData {
  labels: string[];
  datasets: any[];
}

const App: React.FC = () => {
  const [analysisResult, setAnalysisResult] = useState<IAnalysisResult[]>([]);
  const [chartData, setChartData] = useState<IChartData | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
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
        const hourlyCounts: { [eqp: string]: number[] } = {};
        top3Eqps.forEach(eqp => {
          hourlyCounts[eqp] = Array(24).fill(0);
        });

        for (const row of json) {
            const eqp = row['EQP'];
            if (top3Eqps.includes(eqp)) {
                const time = row['Hora'];
                // Handle both string time and Excel's numeric time format
                let hour: number | null = null;
                if (typeof time === 'string') {
                    const hourPart = time.split(':')[0];
                    const parsedHour = parseInt(hourPart, 10);
                    if (!isNaN(parsedHour)) {
                        hour = parsedHour;
                    }
                } else if (typeof time === 'number' && time < 1) { // Excel time is a fraction of a day
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
        chartInstanceRef.current = new Chart(ctx, {
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
     // Cleanup on component unmount
    return () => {
        if (chartInstanceRef.current) {
            chartInstanceRef.current.destroy();
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
        <div className="results-container" role="region" aria-labelledby="results-title">
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
        <div className="chart-container">
            <h2 className="chart-title">Tendência de Atuações por Hora (Top 3)</h2>
            <canvas ref={chartRef} aria-label="Gráfico de linha de atuações por hora" role="img"></canvas>
        </div>
      )}
    </div>
  );
};

export default App;