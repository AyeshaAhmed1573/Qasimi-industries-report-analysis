import React, { useCallback, useRef, useState } from 'react';

const API_BASE = import.meta.env.VITE_API_URL ? `${import.meta.env.VITE_API_URL}/api` : '/api';

export default function App() {
  const [file, setFile] = useState(null);
  const [uploadId, setUploadId] = useState(null);
  const [sheetCount, setSheetCount] = useState(0);
  const [fromDate, setFromDate] = useState('');
  const [toDate, setToDate] = useState('');
  const [dragging, setDragging] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [rows, setRows] = useState(null); // null = no report generated yet
  const [warnings, setWarnings] = useState([]);
  const [error, setError] = useState('');
  const inputRef = useRef(null);

  const resetReport = () => {
    setRows(null);
    setWarnings([]);
    setError('');
  };

  const handleFile = useCallback(async (selectedFile) => {
    if (!selectedFile) return;
    if (!/\.xlsx?$/i.test(selectedFile.name)) {
      setError('Please upload an Excel file (.xlsx or .xls).');
      return;
    }
    setError('');
    setFile(selectedFile);
    setUploadId(null);
    resetReport();
    setUploading(true);

    try {
      const formData = new FormData();
      formData.append('file', selectedFile);
      const res = await fetch(`${API_BASE}/upload`, { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Upload failed.');
      setUploadId(data.uploadId);
      setSheetCount(data.sheetCount);
    } catch (err) {
      setError(err.message);
      setFile(null);
    } finally {
      setUploading(false);
    }
  }, []);

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const dropped = e.dataTransfer.files?.[0];
    handleFile(dropped);
  };

  const onGenerate = async () => {
    if (!uploadId || !fromDate || !toDate) return;
    resetReport();
    setGenerating(true);
    try {
      const res = await fetch(`${API_BASE}/report`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ uploadId, fromDate, toDate }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Could not generate report.');
      setRows(data.rows);
      setWarnings(data.warnings || []);
    } catch (err) {
      setError(err.message);
    } finally {
      setGenerating(false);
    }
  };

  const rangeLabel = fromDate === toDate ? fromDate : `${fromDate}_to_${toDate}`;

  const onDownload = async () => {
    if (!uploadId || !fromDate || !toDate) return;
    setDownloading(true);
    try {
      const res = await fetch(`${API_BASE}/download`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ uploadId, fromDate, toDate }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Download failed.');
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `report-${rangeLabel}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message);
    } finally {
      setDownloading(false);
    }
  };

  const clearFile = () => {
    setFile(null);
    setUploadId(null);
    setSheetCount(0);
    setFromDate('');
    setToDate('');
    resetReport();
    if (inputRef.current) inputRef.current.value = '';
  };

  const canGenerate = uploadId && fromDate && toDate && !generating;

  return (
    <div className="page">
      <div className="header">
        <h1>📅 Daily Sheet Date Report</h1>
        <p>Upload your multi-sheet Excel workbook, pick a date range, and download every matching entry.</p>
      </div>

      {error && <div className="error-banner">{error}</div>}

      <div className="card">
        <div
          className={`dropzone ${dragging ? 'dragging' : ''}`}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
        >
          <div className="icon">📁</div>
          <div className="primary">{uploading ? 'Uploading…' : 'Click or drag your Excel file here'}</div>
          <div className="secondary">.xlsx or .xls — any number of sheets</div>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={(e) => handleFile(e.target.files?.[0])}
          />
        </div>

        {file && (
          <div className="file-chip">
            <span className="name">
              📄 {file.name} {uploadId && `— ${sheetCount} sheet${sheetCount === 1 ? '' : 's'} loaded`}
            </span>
            <button onClick={clearFile}>Remove</button>
          </div>
        )}

        <div className="controls-row">
          <div className="field">
            <label htmlFor="from-date">From date</label>
            <input
              id="from-date"
              type="date"
              value={fromDate}
              max={toDate || undefined}
              onChange={(e) => {
                const val = e.target.value;
                setFromDate(val);
                if (!toDate || val > toDate) setToDate(val);
                resetReport();
              }}
              disabled={!uploadId}
            />
          </div>
          <div className="field">
            <label htmlFor="to-date">To date</label>
            <input
              id="to-date"
              type="date"
              value={toDate}
              min={fromDate || undefined}
              onChange={(e) => {
                setToDate(e.target.value);
                resetReport();
              }}
              disabled={!uploadId}
            />
          </div>
          <button className="primary" disabled={!canGenerate} onClick={onGenerate}>
            {generating ? (
              <>
                <span className="spinner" /> Searching…
              </>
            ) : (
              'Generate Report'
            )}
          </button>
        </div>
      </div>

      {rows !== null && (
        <div className="card">
          <div className="status-row">
            <h2>Matching entries</h2>
            <span className="count-pill">{rows.length} found</span>
          </div>

          {rows.length === 0 ? (
            <div className="empty-state">
              No entries found between {fromDate} and {toDate}. Try a different range.
            </div>
          ) : (
            <>
              <div className="table-scroll">
                <table>
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Sheet</th>
                      <th>Receiver / Issuer</th>
                      <th>Receive Qty</th>
                      <th>Receive Net Wt (kg)</th>
                      <th>Issue Qty</th>
                      <th>Issue Net Wt (kg)</th>
                      <th>Total Receive Qty (Sheet)</th>
                      <th>Total Receive Net Wt (Sheet)</th>
                      <th>Total Issue Qty (Sheet)</th>
                      <th>Total Issue Net Wt (Sheet)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, i) => (
                      <tr key={i}>
                        <td>{r.date}</td>
                        <td>{r.sheet}</td>
                        <td>{r.person}</td>
                        <td>{r.receiveQty ?? '—'}</td>
                        <td>{r.receiveNetWeightKg ?? '—'}</td>
                        <td>{r.issueQty ?? '—'}</td>
                        <td>{r.issueNetWeightKg ?? '—'}</td>
                        <td>{r.totalReceiveQty ?? '—'}</td>
                        <td>{r.totalReceiveNetWeightKg ?? '—'}</td>
                        <td>{r.totalIssueQty ?? '—'}</td>
                        <td>{r.totalIssueNetWeightKg ?? '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="footer-actions">
                <button className="secondary" disabled={downloading} onClick={onDownload}>
                  {downloading ? 'Preparing…' : '⬇ Download Excel'}
                </button>
              </div>
            </>
          )}

          {warnings.length > 0 && (
            <details className="warnings">
              <summary>{warnings.length} sheet warning(s)</summary>
              <ul>
                {warnings.map((w, i) => (
                  <li key={i}>{w}</li>
                ))}
              </ul>
            </details>
          )}
        </div>
      )}
    </div>
  );
}
