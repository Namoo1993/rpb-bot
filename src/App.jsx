import { useState, useEffect, useRef } from 'react'
import './App.css'

const API_BASE = import.meta.env.VITE_API_BASE || ''

function formatDate(iso) {
  if (!iso) return '-'
  const d = new Date(iso)
  return `${d.getFullYear()}년 ${d.getMonth()+1}월 ${d.getDate()}일 ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`
}

function downloadBase64(base64, filename) {
  const link = document.createElement('a')
  link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64
  link.download = filename
  link.click()
}

export default function App() {
  const [file, setFile] = useState(null)
  const [label, setLabel] = useState(() => {
    const now = new Date()
    return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`
  })
  const [status, setStatus] = useState('idle') // idle | processing | done | error
  const [result, setResult] = useState(null)
  const [error, setError]   = useState('')
  const [history, setHistory] = useState([])
  const [dragOver, setDragOver] = useState(false)
  const inputRef = useRef()

  useEffect(() => { loadHistory() }, [])

  async function loadHistory() {
    try {
      const res = await fetch(`${API_BASE}/api/history`)
      const json = await res.json()
      setHistory(json.history || [])
    } catch {}
  }

  function handleFile(f) {
    if (!f) return
    if (!f.name.endsWith('.xlsx') && !f.name.endsWith('.xls')) {
      setError('xlsx 또는 xls 파일만 업로드 가능합니다.')
      return
    }
    setFile(f)
    setError('')
    setStatus('idle')
    setResult(null)
  }

  async function handleSubmit() {
    if (!file) { setError('파일을 먼저 선택해주세요.'); return }
    setStatus('processing')
    setError('')
    setResult(null)
    try {
      const base64 = await new Promise((res, rej) => {
        const reader = new FileReader()
        reader.onload = e => res(e.target.result.split(',')[1])
        reader.onerror = rej
        reader.readAsDataURL(file)
      })
      const response = await fetch(`${API_BASE}/api/process`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileData: base64, label }),
      })
      const json = await response.json()
      if (!response.ok) throw new Error(json.error || '처리 실패')
      setResult(json)
      setStatus('done')
      // 즉시 다운로드
      downloadBase64(json.fileBase64, json.filename)
      // 히스토리 새로고침
      loadHistory()
    } catch (e) {
      setError(e.message)
      setStatus('error')
    }
  }

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <span className="logo-mark">RPB</span>
          <div className="header-title">
            <h1>분석표 자동 업데이트</h1>
            <p>RAW DATA를 업로드하면 RPB 분석표를 자동으로 채워드립니다</p>
          </div>
        </div>
      </header>

      <main className="main">
        {/* 업로드 패널 */}
        <section className="card upload-card">
          <h2 className="section-title">① RAW DATA 업로드</h2>

          <div
            className={`drop-zone ${dragOver ? 'drag-over' : ''} ${file ? 'has-file' : ''}`}
            onClick={() => inputRef.current.click()}
            onDragOver={e => { e.preventDefault(); setDragOver(true) }}
            onDragLeave={() => setDragOver(false)}
            onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]) }}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls"
              style={{ display:'none' }}
              onChange={e => handleFile(e.target.files[0])}
            />
            {file ? (
              <div className="file-info">
                <span className="file-icon">📊</span>
                <span className="file-name">{file.name}</span>
                <span className="file-size">{(file.size/1024).toFixed(1)} KB</span>
                <button className="btn-clear" onClick={e => { e.stopPropagation(); setFile(null); setStatus('idle'); setResult(null) }}>✕</button>
              </div>
            ) : (
              <div className="drop-hint">
                <span className="drop-icon">⬆</span>
                <span>클릭하거나 파일을 드래그하세요</span>
                <span className="drop-sub">.xlsx / .xls</span>
              </div>
            )}
          </div>

          <div className="label-row">
            <label className="label-text">기간 레이블</label>
            <input
              className="label-input"
              type="text"
              value={label}
              onChange={e => setLabel(e.target.value)}
              placeholder="예: 2026-06"
            />
          </div>

          {error && <p className="error-msg">⚠ {error}</p>}

          <button
            className={`btn-run ${status === 'processing' ? 'running' : ''}`}
            onClick={handleSubmit}
            disabled={status === 'processing' || !file}
          >
            {status === 'processing' ? (
              <><span className="spinner" />처리 중...</>
            ) : '분석표 생성 →'}
          </button>
        </section>

        {/* 결과 패널 */}
        {status === 'done' && result && (
          <section className="card result-card">
            <h2 className="section-title">② 결과</h2>
            <div className="result-grid">
              <div className="result-stat">
                <span className="stat-value">{result.totalTeu?.toLocaleString()}</span>
                <span className="stat-label">집계 TEU</span>
              </div>
              <div className="result-stat">
                <span className="stat-value">{result.filename?.split('_').slice(-1)[0].replace('.xlsx','')}</span>
                <span className="stat-label">생성 시각</span>
              </div>
            </div>
            <div className="result-actions">
              <button className="btn-dl primary" onClick={() => downloadBase64(result.fileBase64, result.filename)}>
                ↓ 다시 다운로드
              </button>
              <a className="btn-dl secondary" href={result.url} target="_blank" rel="noreferrer">
                🔗 링크 복사용
              </a>
            </div>
          </section>
        )}

        {/* 히스토리 */}
        <section className="card history-card">
          <div className="history-header">
            <h2 className="section-title">③ 히스토리</h2>
            <button className="btn-refresh" onClick={loadHistory}>↺ 새로고침</button>
          </div>
          {history.length === 0 ? (
            <p className="empty-history">아직 생성된 파일이 없습니다.</p>
          ) : (
            <table className="history-table">
              <thead>
                <tr><th>기간</th><th>TEU</th><th>생성일시</th><th>파일</th></tr>
              </thead>
              <tbody>
                {history.map(h => (
                  <tr key={h.id}>
                    <td><span className="badge">{h.label || '-'}</span></td>
                    <td className="teu-cell">{h.total_teu?.toLocaleString()}</td>
                    <td className="date-cell">{formatDate(h.created_at)}</td>
                    <td>
                      <a className="dl-link" href={h.url} target="_blank" rel="noreferrer" download>
                        ↓ 다운로드
                      </a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </section>
      </main>
    </div>
  )
}
