import React, { useState, useCallback } from 'react';
import { createRoot } from 'react-dom/client';

// Make SheetJS library available in the scope, as it's loaded from a script tag
declare const XLSX: any;

interface IAnalysisResult {
  eqp: string;
  count: number;
}

const App: React.FC = () => {
  const [analysisResult, setAnalysisResult] = useState<IAnalysisResult[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);


  const processFile = useCallback((file: File) => {
    setIsLoading(true);
    setError(null);
    setAnalysisResult([]);
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

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      if (file.name.endsWith('.xlsx')) {
        processFile(file);
      } else {
        setError("Por favor, selecione um arquivo .xlsx");
      }
    }
     // Reset file input to allow re-uploading the same file
    event.target.value = '';
  };

  return (
    <div className="container">
      <h1 className="title">Analisador de Atuações de Equipamentos</h1>
      <p className="description">
        Importe uma planilha (.xlsx) com as colunas "EQP", "Data" e "Hora" para analisar a quantidade de atuações dos 10 principais equipamentos.
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
    </div>
  );
};

const container = document.getElementById('root');
const root = createRoot(container!);
root.render(<App />);
