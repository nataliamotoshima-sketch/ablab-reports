import { useState, useRef } from 'react'
import Head from 'next/head'

const PURPLE = '#7B2FBE'
const PURPLE_LIGHT = '#9B4DCA'
const BG = '#F5F0E8'

export default function Home() {
  const [marca, setMarca] = useState('')
  const [periodo, setPeriodo] = useState('')
  const [meta, setMeta] = useState('')
  const [mesAnt, setMesAnt] = useState('')
  const [histWhats, setHistWhats] = useState('')
  const [histTuo, setHistTuo] = useState('')
  const [histTel, setHistTel] = useState('')
  const [contexto, setContexto] = useState('')
  const [proximos, setProximos] = useState('')

  const [fileGoogle, setFileGoogle] = useState(null)
  const [fileMeta, setFileMeta] = useState(null)
  const [fileGA4, setFileGA4] = useState(null)

  const [status, setStatus] = useState(null) // null | 'loading' | 'done' | 'error'
  const [statusMsg, setStatusMsg] = useState('')
  const [downloadUrl, setDownloadUrl] = useState(null)

  const refGoogle = useRef()
  const refMeta = useRef()
  const refGA4 = useRef()

  async function gerar() {
    if (!fileGoogle || !fileMeta || !fileGA4) {
      setStatus('error')
      setStatusMsg('Por favor, faça upload dos 3 arquivos (Google Ads, Meta Ads e GA4).')
      return
    }
    if (!marca || !periodo || !meta) {
      setStatus('error')
      setStatusMsg('Preencha ao menos: nome da marca, período e meta de conversões.')
      return
    }

    setStatus('loading')
    setStatusMsg('Processando dados e gerando textos com IA...')
    setDownloadUrl(null)

    const form = new FormData()
    form.append('google', fileGoogle)
    form.append('meta', fileMeta)
    form.append('ga4', fileGA4)
    form.append('marca', marca)
    form.append('periodo', periodo)
    form.append('meta_conv', meta)
    form.append('mes_ant', mesAnt || 'mês anterior')
    form.append('hist_whats', histWhats || '0')
    form.append('hist_tuo', histTuo || '0')
    form.append('hist_tel', histTel || '0')
    form.append('contexto', contexto)
    form.append('proximos', proximos)

    try {
      const res = await fetch('/api/gerar', { method: 'POST', body: form })
      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.message || 'Erro ao gerar report')
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      setDownloadUrl(url)
      setStatus('done')
      setStatusMsg('Report gerado com sucesso!')
    } catch (e) {
      setStatus('error')
      setStatusMsg(e.message)
    }
  }

  function FileZone({ label, icon, accept, file, setFile, inputRef, platform }) {
    return (
      <div
        onClick={() => inputRef.current.click()}
        style={{
          border: `1.5px dashed ${file ? PURPLE : '#C4B8D8'}`,
          borderRadius: 12,
          padding: '20px 16px',
          textAlign: 'center',
          cursor: 'pointer',
          background: file ? '#F3EEFF' : '#FDFBF7',
          transition: 'all 0.2s',
        }}
      >
        <div style={{ fontSize: 28, marginBottom: 6 }}>{icon}</div>
        <div style={{ fontSize: 13, fontWeight: 600, color: '#3A2060', marginBottom: 4 }}>{platform}</div>
        <div style={{ fontSize: 11, color: '#888', marginBottom: 8 }}>{label}</div>
        {file ? (
          <div style={{ fontSize: 11, color: PURPLE, fontWeight: 600 }}>✓ {file.name}</div>
        ) : (
          <div style={{ fontSize: 11, color: '#AAA' }}>clique para selecionar</div>
        )}
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          style={{ display: 'none' }}
          onChange={e => setFile(e.target.files[0] || null)}
        />
      </div>
    )
  }

  return (
    <>
      <Head>
        <title>Gerador de Reports — a/b lab</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" />
      </Head>

      <div style={{ minHeight: '100vh', background: BG, fontFamily: "'Inter', sans-serif" }}>

        {/* Header */}
        <div style={{ background: '#3A2060', padding: '16px 32px', display: 'flex', alignItems: 'center', gap: 16 }}>
          <div style={{ background: PURPLE, borderRadius: 8, padding: '6px 14px', fontSize: 18, fontWeight: 800, color: '#fff', letterSpacing: -0.5 }}>
            a/b lab
          </div>
          <div style={{ color: '#C4A8FF', fontSize: 15, fontWeight: 500 }}>Gerador de Report Mensal</div>
        </div>

        <div style={{ maxWidth: 780, margin: '0 auto', padding: '40px 24px' }}>

          {/* Section 1 */}
          <Section title="1. Dados do relatório">
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
              <Field label="Nome da marca" value={marca} onChange={setMarca} placeholder="Ex: Pro Matre" />
              <Field label="Período" value={periodo} onChange={setPeriodo} placeholder="Ex: 01 a 31 de maio de 2026" />
              <Field label="Meta de conversões (Google)" value={meta} onChange={setMeta} placeholder="Ex: 2500" type="number" />
              <Field label="Mês anterior (para comparativo)" value={mesAnt} onChange={setMesAnt} placeholder="Ex: abril" />
            </div>
            <div style={{ marginTop: 16 }}>
              <div style={{ fontSize: 13, color: '#666', fontWeight: 500, marginBottom: 8 }}>
                Botões de agendamento — mês anterior (Whats / Tuotempo / Telefone)
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12 }}>
                <Field label="WhatsApp" value={histWhats} onChange={setHistWhats} placeholder="Ex: 1078" type="number" />
                <Field label="Tuotempo" value={histTuo} onChange={setHistTuo} placeholder="Ex: 1025" type="number" />
                <Field label="Telefone" value={histTel} onChange={setHistTel} placeholder="Ex: 163" type="number" />
              </div>
            </div>
          </Section>

          {/* Section 2 */}
          <Section title="2. Upload dos arquivos">
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 16 }}>
              <FileZone
                label="Performance por campanha (.csv)"
                icon="🔵"
                accept=".csv"
                file={fileGoogle}
                setFile={setFileGoogle}
                inputRef={refGoogle}
                platform="Google Ads CSV"
              />
              <FileZone
                label="Campanhas do período (.xlsx ou .csv)"
                icon="🟣"
                accept=".xlsx,.csv"
                file={fileMeta}
                setFile={setFileMeta}
                inputRef={refMeta}
                platform="Meta Ads XLSX"
              />
              <FileZone
                label="Conversões x Campanhas (.csv)"
                icon="🟢"
                accept=".csv"
                file={fileGA4}
                setFile={setFileGA4}
                inputRef={refGA4}
                platform="GA4 CSV"
              />
            </div>
            <div style={{ marginTop: 10, fontSize: 11, color: '#999' }}>
              Os arquivos são processados no servidor e não são armazenados. Mesmos formatos de exportação padrão das plataformas.
            </div>
          </Section>

          {/* Section 3 */}
          <Section title="3. Contexto adicional (opcional)">
            <TextArea
              label="Acontecimentos do mês (feriados, mudanças de campanha, eventos...)"
              value={contexto}
              onChange={setContexto}
              placeholder="Ex: Semana 1 impactada pelo feriado do Dia do Trabalho. Novo criativo lançado na semana 2..."
            />
            <TextArea
              label="Próximos passos planejados"
              value={proximos}
              onChange={setProximos}
              placeholder="Ex: Atualização de textos das campanhas, novo teste A/B no Meta, expansão de budget na Institucional..."
            />
          </Section>

          {/* Generate button */}
          <button
            onClick={gerar}
            disabled={status === 'loading'}
            style={{
              width: '100%',
              padding: '16px',
              background: status === 'loading' ? '#9B4DCA' : PURPLE,
              color: '#fff',
              border: 'none',
              borderRadius: 12,
              fontSize: 16,
              fontWeight: 700,
              cursor: status === 'loading' ? 'not-allowed' : 'pointer',
              letterSpacing: 0.2,
              transition: 'background 0.2s',
            }}
          >
            {status === 'loading' ? '⏳ Gerando report...' : '✨ Gerar Report em PowerPoint'}
          </button>

          {/* Status */}
          {status && (
            <div style={{
              marginTop: 16,
              padding: '14px 18px',
              borderRadius: 10,
              background: status === 'done' ? '#EAF6EF' : status === 'error' ? '#FDECEC' : '#EEE8FA',
              color: status === 'done' ? '#1A6B3C' : status === 'error' ? '#9B1C1C' : '#3A2060',
              fontSize: 14,
              fontWeight: 500,
            }}>
              {status === 'loading' && <span>⏳ </span>}
              {status === 'done' && <span>✅ </span>}
              {status === 'error' && <span>❌ </span>}
              {statusMsg}
            </div>
          )}

          {/* Download */}
          {downloadUrl && (
            <a
              href={downloadUrl}
              download={`Report_${marca.replace(/\s+/g, '_')}_${periodo.replace(/\s+/g, '_')}.pptx`}
              style={{
                display: 'block',
                marginTop: 16,
                padding: '14px',
                background: '#3A2060',
                color: '#fff',
                borderRadius: 12,
                textAlign: 'center',
                fontSize: 15,
                fontWeight: 700,
                textDecoration: 'none',
              }}
            >
              ⬇️ Baixar PowerPoint
            </a>
          )}

        </div>
      </div>
    </>
  )
}

function Section({ title, children }) {
  return (
    <div style={{ marginBottom: 32 }}>
      <div style={{ fontSize: 15, fontWeight: 700, color: '#3A2060', marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
        <span style={{ width: 4, height: 18, background: PURPLE, borderRadius: 2, display: 'inline-block' }} />
        {title}
      </div>
      {children}
    </div>
  )
}

function Field({ label, value, onChange, placeholder, type = 'text' }) {
  return (
    <div>
      <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#555', marginBottom: 6 }}>{label}</label>
      <input
        type={type}
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        style={{
          width: '100%',
          padding: '9px 12px',
          border: '1px solid #D4C8E8',
          borderRadius: 8,
          fontSize: 14,
          color: '#222',
          background: '#FDFBF7',
          outline: 'none',
          boxSizing: 'border-box',
        }}
      />
    </div>
  )
}

function TextArea({ label, value, onChange, placeholder }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#555', marginBottom: 6 }}>{label}</label>
      <textarea
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        rows={3}
        style={{
          width: '100%',
          padding: '9px 12px',
          border: '1px solid #D4C8E8',
          borderRadius: 8,
          fontSize: 14,
          color: '#222',
          background: '#FDFBF7',
          resize: 'vertical',
          outline: 'none',
          boxSizing: 'border-box',
          fontFamily: 'inherit',
        }}
      />
    </div>
  )
}
