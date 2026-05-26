import formidable from 'formidable'
import fs from 'fs'
import path from 'path'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import Anthropic from 'anthropic'

export const config = { api: { bodyParser: false } }

// ─── Parse helpers ───────────────────────────────────────────
function parseBR(v) {
  if (v === undefined || v === null || v === '') return 0
  return parseFloat(String(v).replace(/\./g, '').replace(',', '.').replace('%', '').trim()) || 0
}

function parseGoogleCSV(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8')
  const lines = raw.split('\n')
  const hi = lines.findIndex(l => l.startsWith('Campanha,') || l.startsWith('Campaign,'))
  if (hi < 0) throw new Error('Formato Google Ads CSV não reconhecido')
  const headers = lines[hi].split(',').map(h => h.trim())
  const rows = []
  for (let i = hi + 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue
    const cols = lines[i].split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(c => c.replace(/"/g, '').trim())
    const obj = {}
    headers.forEach((h, j) => obj[h] = cols[j] || '')
    if (obj['Campanha'] || obj['Campaign']) rows.push(obj)
  }
  return rows
}

function parseMetaXLSX(filePath) {
  const wb = XLSX.readFile(filePath)
  const ws = wb.Sheets[wb.SheetNames[0]]
  return XLSX.utils.sheet_to_json(ws, { defval: '' })
}

function parseGA4CSV(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8')
  const lines = raw.split('\n')
  const hi = lines.findIndex(l => l.startsWith('Nome do evento') || l.startsWith('Event name'))
  if (hi < 0) throw new Error('Formato GA4 CSV não reconhecido')
  // Line hi+1 is sub-header, hi+2 is totals row
  const totalLine = lines[hi + 2]?.trim().split(',') || []
  return {
    tuotempo: parseInt(totalLine[1]) || 0,
    whatsapp: (parseInt(totalLine[2]) || 0) + (parseInt(totalLine[4]) || 0),
    telefone: (parseInt(totalLine[3]) || 0) + (parseInt(totalLine[5]) || 0),
    total: parseInt(totalLine[6]) || 0,
    rawLines: lines.slice(hi + 2).filter(l => l.trim()),
  }
}

function calcTotals(googleRows, metaRows, ga4) {
  const g = { custo: 0, impr: 0, cli: 0, conv: 0 }
  googleRows.forEach(r => {
    g.custo += parseBR(r['Custo'] || r['Cost'])
    g.impr  += parseBR(r['Impr.'] || r['Impressions'])
    g.cli   += parseBR(r['Cliques'] || r['Clicks'])
    g.conv  += parseBR(r['Conversões'] || r['Conversions'])
  })
  g.ctr = g.impr ? (g.cli / g.impr * 100) : 0
  g.cpc = g.cli  ? g.custo / g.cli : 0
  g.cpa = g.conv ? g.custo / g.conv : 0
  g.txc = g.cli  ? (g.conv / g.cli * 100) : 0

  const m = { custo: 0, impr: 0, cli: 0 }
  metaRows.forEach(r => {
    m.custo += parseFloat(String(r['Amount spent (BRL)'] || r['Valor usado (BRL)'] || 0).replace(',', '')) || 0
    m.impr  += parseFloat(String(r['Impressions'] || 0).replace(',', '')) || 0
    m.cli   += parseFloat(String(r['Link clicks'] || 0).replace(',', '')) || 0
  })
  m.ctr = m.impr ? (m.cli / m.impr * 100) : 0
  m.cpc = m.cli  ? m.custo / m.cli : 0

  return { g, m }
}

// ─── Generate texts via Claude ────────────────────────────────
async function generateTexts({ g, m, ga4, marca, periodo, meta, mesAnt, contexto, proximos, histWhats, histTuo, histTel }) {
  const client = new Anthropic()
  const pctMeta = meta ? Math.round(g.conv / meta * 100) : null
  const brl = v => `R$ ${v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  const num = v => Math.round(v).toLocaleString('pt-BR')

  const prompt = `Você é analista sênior de mídia digital da agência a/b lab. Gere os textos para um report mensal de mídia paga no estilo da agência: direto, analítico, com dados embasando cada afirmação.

DADOS DO MÊS:
- Marca: ${marca}
- Período: ${periodo}
- Meta de conversões: ${num(meta)}
- Conversões Google: ${num(g.conv)}${pctMeta ? ` (${pctMeta}% da meta)` : ''}

GOOGLE ADS:
- Investimento: ${brl(g.custo)} | Impressões: ${num(g.impr)} | CTR: ${g.ctr.toFixed(1)}%
- Cliques: ${num(g.cli)} | CPC: ${brl(g.cpc)} | Tx. Conversão: ${g.txc.toFixed(1)}% | CPA: ${brl(g.cpa)}

META ADS:
- Investimento: ${brl(m.custo)} | Impressões: ${num(m.impr)} | Cliques: ${num(m.cli)}
- CTR: ${m.ctr.toFixed(2)}% | CPC: ${brl(m.cpc)}

BOTÕES DE AGENDAMENTO (GA4):
- Tuotempo: ${num(ga4.tuotempo)} | WhatsApp: ${num(ga4.whatsapp)} | Telefone: ${num(ga4.telefone)} | Total: ${num(ga4.total)}

HISTÓRICO ${mesAnt.toUpperCase()}:
- Tuotempo: ${num(histTuo)} | WhatsApp: ${num(histWhats)} | Telefone: ${num(histTel)}

VARIAÇÕES vs ${mesAnt}:
- Tuotempo: ${histTuo ? Math.round((ga4.tuotempo - histTuo) / histTuo * 100) + '%' : 'n/d'}
- WhatsApp: ${histWhats ? Math.round((ga4.whatsapp - histWhats) / histWhats * 100) + '%' : 'n/d'}
- Telefone: ${histTel ? Math.round((ga4.telefone - histTel) / histTel * 100) + '%' : 'n/d'}

CONTEXTO DO MÊS: ${contexto || 'Não informado'}
PRÓXIMOS PASSOS PLANEJADOS: ${proximos || 'Não informado'}

Retorne SOMENTE um JSON válido, sem markdown, com estas chaves:
{
  "intro_p1": "Parágrafo 1 do Big Numbers (contexto estratégico do mês, 2-3 frases)",
  "intro_p2": "Parágrafo 2 do Big Numbers (resultado geral com números, 2-3 frases)",
  "dados_perf": "Texto do slide Dados de Performance (1 parágrafo sobre trajetória de conversões)",
  "botoes_p1": "Parágrafo 1 dos cliques nos botões (análise das variações, destaque positivo)",
  "botoes_p2": "Parágrafo 2 dos botões (contexto e implicações estratégicas)",
  "campanha_p1": "Parágrafo 1 dos botões por campanha (campanha principal)",
  "campanha_p2": "Parágrafo 2 dos botões por campanha (campanhas complementares)",
  "sem1": "Análise semana 1 (2-3 frases)",
  "sem2": "Análise semana 2 (2-3 frases)",
  "sem3": "Análise semana 3 (2-3 frases)",
  "sem4": "Análise semana 4 (2-3 frases)",
  "sem5": "Análise semana 5 ou encerramento (2-3 frases)",
  "funil_p1": "Texto funil parágrafo 1 (qualidade do tráfego, 2-3 frases)",
  "funil_p2": "Texto funil parágrafo 2 (conversões vs meta, 2-3 frases)",
  "termos_busca1": "Texto Buscas Institucionais (2-3 frases)",
  "termos_busca2": "Texto Intenção de Agendamento (2-3 frases)",
  "termos_busca3": "Texto Busca por Especialidades (2-3 frases)",
  "projecao": "Texto projeção próximo mês (2-3 frases)",
  "conclusao_p1": "Conclusão parágrafo 1 (resultado geral do mês com números)",
  "conclusao_p2": "Conclusão parágrafo 2 (botões e canais)",
  "conclusao_p3": "Conclusão parágrafo 3 (Meta Ads e papel estratégico)",
  "prox1": "Próximo passo 1 (1 frase)",
  "prox2": "Próximo passo 2 (1 frase)",
  "prox3": "Próximo passo 3 (1 frase)"
}`

  const response = await client.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2000,
    messages: [{ role: 'user', content: prompt }],
  })

  const raw = response.content.map(c => c.text || '').join('').replace(/```json|```/g, '').trim()
  return JSON.parse(raw)
}

// ─── PPTX XML editing ────────────────────────────────────────
function replaceXmlText(xml, oldText, newText) {
  return xml.replace(oldText, newText)
}

function buildTableRow(zip, slideXml, oldVal, newVal) {
  return slideXml.split(oldVal).join(newVal)
}

async function buildPPTX(params) {
  const { g, m, ga4, marca, periodo, meta, mesAnt, histWhats, histTuo, histTel, texts } = params
  const pctMeta = meta ? Math.round(g.conv / meta * 100) : null

  const brl = v => `R$ ${v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  const num = v => Math.round(v).toLocaleString('pt-BR')
  const pct = v => `${v.toFixed(1)}%`

  // Load template
  const templatePath = path.join(process.cwd(), 'public', 'template.pptx')
  const templateBuffer = fs.readFileSync(templatePath)
  const zip = await JSZip.loadAsync(templateBuffer)

  async function editSlide(n, fn) {
    const key = `ppt/slides/slide${n}.xml`
    let xml = await zip.file(key).async('string')
    xml = fn(xml)
    zip.file(key, xml)
  }

  function r(xml, from, to) { return xml.split(from).join(to) }

  // ── Slide 1: Capa
  await editSlide(1, xml => r(xml, ' : 01 a 31 de abril', ` : ${periodo}`))

  // ── Slide 3: Big Numbers
  await editSlide(3, xml => {
    xml = r(xml, ' Abril', ' Maio')
    xml = r(xml, 'Em abril, seguimos consolidando a estratégia de campanhas orientadas para conversão, implementada em março. Como esperado nesse tipo de otimização, o algoritmo ganha mais eficiência com o tempo, aprimorando a entrega para públicos com maior propensão à ação e fortalecendo a qualidade das conversões ao longo do período.', texts.intro_p1)
    xml = r(xml, 'Mesmo diante de um mês impactado por feriados prolongados e por uma última semana antecedendo o feriado de 1º de maio, fatores que costumam influenciar o comportamento do usuário e reduzir o volume de interações, encerramos abril com 2.288 conversões, atingindo 121% da meta mensal (1.895);  reforçando a consistência da estratégia e a eficiência da mídia na geração de resultados.', texts.intro_p2)
    // Google row
    xml = r(xml, 'R$ 6.953,84', brl(g.custo))
    xml = r(xml, '>42.710<', `>${num(g.impr)}<`)
    xml = r(xml, '>15,1%<', `>${g.ctr.toFixed(1)}%<`)
    xml = r(xml, '>R$ 1,08<', `>${brl(g.cpc)}<`)
    xml = r(xml, '>6.434<', `>${num(g.cli)}<`)
    xml = r(xml, '>33%<', `>${pct(g.txc)}<`)
    xml = r(xml, '>2.275<', `>${num(g.conv)}<`)
    xml = r(xml, '>R$ 3,06<', `>${brl(g.cpa)}<`)
    // Meta row
    xml = r(xml, 'R$ 1.773,62', brl(m.custo))
    xml = r(xml, '>775.877<', `>${num(m.impr)}<`)
    xml = r(xml, '>1,3%<', `>${m.ctr.toFixed(2)}%<`)
    xml = r(xml, '>R$ 0,17<', `>${brl(m.cpc)}<`)
    xml = r(xml, '>10.322<', `>${num(m.cli)}<`)
    xml = r(xml, '>1%<', '>—<')
    xml = r(xml, '>13<', '>—<')
    xml = r(xml, '>R$ 136,43<', '>—<')
    // Total row
    xml = r(xml, 'R$ 8.727,46', brl(g.custo + m.custo))
    xml = r(xml, '>818.587<', `>${num(g.impr + m.impr)}<`)
    xml = r(xml, '>2,05%<', `>${((g.cli + m.cli) / (g.impr + m.impr) * 100).toFixed(1)}%<`)
    xml = r(xml, '>R$ 0,52<', `>${brl((g.custo + m.custo) / (g.cli + m.cli))}<`)
    xml = r(xml, '>16.756<', `>${num(g.cli + m.cli)}<`)
    xml = r(xml, '>26,22%<', `>${pct(g.txc)}<`)
    xml = r(xml, '>2.288<', `>${num(g.conv)}<`)
    xml = r(xml, '>R$ 3,81<', `>${brl(g.cpa)}<`)
    return xml
  })

  // ── Slide 4: Dados de Performance
  await editSlide(4, xml => {
    xml = r(xml, '>43%<', `>${pctMeta ? pctMeta + '%' : '—'}<`)
    xml = r(xml, ' Desde janeiro, acompanhamos uma projeção contínua de crescimento no volume de conversões, com metas mensais estruturadas para sustentar essa evolução de forma consistente. Os resultados de abril reforçam essa trajetória, mostrando que o planejamento vem sendo executado com eficiência e aderência às expectativas de crescimento.', ' ' + texts.dados_perf)
    return xml
  })

  // ── Slide 5: Cliques nos botões (tabela histórico)
  await editSlide(5, xml => {
    xml = r(xml, 'Em abril, observamos uma leve retração no volume de conversões via WhatsApp, movimento concentrado principalmente na semana da Páscoa, período em que feriados costumam impactar diretamente o comportamento e a disponibilidade dos usuários para interação.', texts.botoes_p1)
    xml = r(xml, 'Ainda assim, o cenário geral se manteve positivo, com crescimento expressivo nos canais de Tuotempo e Telefone. Esse avanço foi impulsionado pelos ajustes realizados nos botões da landing page no final de março, conforme recomendação estratégica, melhorando a visibilidade e usabilidade desses pontos de contato e contribuindo para o aumento das interações nesses canais.', texts.botoes_p2)
    // Row 1: Fevereiro -> 2 months ago (mesAnt previous)
    xml = r(xml, '>Fevereiro<', '>Mar&#xE7;o<')
    xml = r(xml, '>975<', '>1.143<')
    xml = r(xml, '>736<', '>523<')
    xml = r(xml, '>96<', '>146<')
    // Row 2: Março -> mesAnt
    const mesAntCap = mesAnt.charAt(0).toUpperCase() + mesAnt.slice(1)
    xml = r(xml, '>Mar&#xE7;o<', `>${mesAntCap}<`)
    xml = r(xml, '>1.143<', `>${num(histWhats)}<`)
    xml = r(xml, '>523<', `>${num(histTuo)}<`)
    xml = r(xml, '>146<', `>${num(histTel)}<`)
    // Row 3: Abril -> current month
    const mesCurrent = periodo.split(' ').find(w => ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'].includes(w.toLowerCase())) || 'Mês atual'
    const mesCurrentCap = mesCurrent.charAt(0).toUpperCase() + mesCurrent.slice(1)
    xml = r(xml, '>Abril<', `>${mesCurrentCap}<`)
    xml = r(xml, '>1.078<', `>${num(ga4.whatsapp)}<`)
    xml = r(xml, '>1.025<', `>${num(ga4.tuotempo)}<`)
    xml = r(xml, '>163<', `>${num(ga4.telefone)}<`)
    return xml
  })

  // ── Slide 6: Botões por campanha
  await editSlide(6, xml => {
    xml = r(xml, 'A campanha Institucional segue como principal responsável pelo volume de conversões, com destaque para o canal Tuotempo, que concentrou a maior parte dos agendamentos no período. Esse comportamento reforça a força da campanha na captura de demanda mais qualificada e na geração de ações de maior intenção.', texts.campanha_p1)
    xml = r(xml, 'As campanhas de Exames e Pré-Natal mantiveram participação complementar na estratégia, contribuindo para diversificar os pontos de entrada e fortalecer a presença da marca em frentes específicas de serviço. Mesmo com menor volume, esses resultados apoiam a construção de relevância e sustentação do ecossistema de conversão como um todo.', texts.campanha_p2)
    return xml
  })

  // ── Slide 7: Performance semanal
  await editSlide(7, xml => {
    xml = r(xml, ' A primeira semana de abril foi marcada pelos ajustes de orçamentoa, etapa importante para otimizar a entrega ao longo do mês. Além disso, o período coincidiu com o feriado de Páscoa, impactando o comportamento dos usuários e reduzindo naturalmente o volume de interações.', ' ' + texts.sem1)
    xml = r(xml, 'Na segunda semana de abril, observamos uma retomada da demanda após o período de feriado, marcada pelo aumento nos picos de cliques no botão de agendamento, voltando a ultrapassar a marca de 100 interações, além do ganho de tração impulsionado pelas primeiras ativações e otimizações de mídia do mês.', texts.sem2)
    xml = r(xml, ' A terceira semana de abril foi o melhor período do mês em performance e aproveitamento de verba. Com menos impactos sazonais em comparação à primeira e à quarta semana, concentramos maior estabilidade de demanda e melhor eficiência das campanhas.', ' ' + texts.sem3)
    xml = r(xml, 'Na quarta semana, o feriado de Tiradentes impactou novamente o volume de demanda. Ainda assim, logo após o período, observamos uma retomada das interações, mostrando recuperação rápida e manutenção do interesse dos usuários.', texts.sem4)
    xml = r(xml, 'A quinta semana antecedeu o feriado de 1º de maio e contou com apenas quatro dias de veiculação, cenário que naturalmente reduz o volume de demanda. Ainda assim, mantivemos a estabilidade da entrega e sustentação dos resultados no fechamento do mês.', texts.sem5)
    return xml
  })

  // ── Slide 8: Funil
  await editSlide(8, xml => {
    xml = r(xml, 'O funil de conversão em abril manteve consistência ao longo do período, com bom volume de alcance e evolução qualificada até os botões de agendamento, refletindo a maturação da estratégia de campanhas orientadas para conversão e o ganho de eficiência das otimizações implementadas desde março.', texts.funil_p1)
    xml = r(xml, 'Mesmo em um mês impactado por feriados e oscilações naturais de demanda, encerramos o período com 2.288 conversões, superando a meta mensal em 21%, o que reforça a eficiência da mídia e a capacidade de sustentar resultados mesmo em cenários de sazonalidade.', texts.funil_p2)
    return xml
  })

  // ── Slide 9: Termos buscados
  await editSlide(9, xml => {
    xml = r(xml, 'Termos mais buscados em Abril', `Termos mais buscados em ${periodo.split(' ').find(w => w.length > 4) || 'Maio'}`)
    xml = r(xml, 'Os termos relacionados à marca seguiram liderando o volume de buscas, com destaque para &#x201C;Pro Matre&#x201D; e suas variações, reforçando o alto reconhecimento da marca e a forte intenção de usuários já familiarizados com a  Pro Matre.', texts.termos_busca1)
    xml = r(xml, 'O termo &#x201C;Pro Matre agendamento&#x201D; se destacou pela alta taxa de clique, evidenciando um público em estágio avançado de decisão e forte propensão à conversão, alinhado ao objetivo das campanhas.', texts.termos_busca2)
    xml = r(xml, 'Termos como &#x201C;ecocardiograma fetal&#x201D; e &#x201C;ultrassom morfológico 1 trimestre&#x201D; reforçam o interesse por serviços específicos, mostrando oportunidades para ampliar a cobertura em campanhas de exames e especialidades.', texts.termos_busca3)
    return xml
  })

  // ── Slide 10: Projeção
  await editSlide(10, xml => {
    xml = r(xml, 'Projeção de Resultados para Maio', `Projeção de Resultados para próximo mês`)
    xml = r(xml, 'A projeção acompanha a curva de amadurecimento das campanhas e mantém uma expectativa de crescimento sustentável, baseada no ganho de eficiência observado desde a mudança de objetivo implementada em março.', texts.projecao)
    return xml
  })

  // ── Slide 11: Conclusão
  await editSlide(11, xml => {
    xml = r(xml, 'Abril foi um mês marcado por desafios sazonais, com feriados que impactaram diretamente o comportamento da demanda ao longo do período. Ainda assim, a estratégia manteve consistência e superou a meta mensal, encerrando o mês com 121% da meta de conversão.', texts.conclusao_p1)
    xml = r(xml, 'A mudança de objetivo para conversão, implementada em março, segue em processo de amadurecimento e ganho de tração, refletindo em maior eficiência na entrega e evolução do volume de agendamentos.', texts.conclusao_p2)
    xml = r(xml, 'Observamos crescimento relevante em canais estratégicos como Tuotempo e Telefone, reforçando a diversificação dos pontos de conversão e sustentando os resultados mesmo diante de oscilações naturais de demanda.', texts.conclusao_p3)
    xml = r(xml, 'Atualização de textos e descrições das campanhas, buscando renovação da comunicação e ganho de relevância.', texts.prox1)
    xml = r(xml, 'Duplicação das campanhas no Meta Ads  para testar uma nova fase de aprendizado e avaliar potencial ganho de CTR e eficiência de entrega.', texts.prox2)
    xml = r(xml, 'Continuidade no processo de otimização e acompanhamento dos padrões de comportamento dos usuários, visando identificar oportunidades estratégicas para futuras ações.', texts.prox3)
    return xml
  })

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

// ─── Main handler ─────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end()

  const form = formidable({ uploadDir: '/tmp', keepExtensions: true, maxFileSize: 20 * 1024 * 1024 })

  form.parse(req, async (err, fields, files) => {
    if (err) return res.status(400).json({ message: 'Erro ao processar arquivos: ' + err.message })

    try {
      const get = k => Array.isArray(fields[k]) ? fields[k][0] : fields[k]
      const getFile = k => Array.isArray(files[k]) ? files[k][0] : files[k]

      const marca     = get('marca') || 'Marca'
      const periodo   = get('periodo') || ''
      const meta      = parseFloat(get('meta_conv')) || 0
      const mesAnt    = get('mes_ant') || 'mês anterior'
      const histWhats = parseFloat(get('hist_whats')) || 0
      const histTuo   = parseFloat(get('hist_tuo')) || 0
      const histTel   = parseFloat(get('hist_tel')) || 0
      const contexto  = get('contexto') || ''
      const proximos  = get('proximos') || ''

      const googleRows = parseGoogleCSV(getFile('google').filepath)
      const metaRows   = parseMetaXLSX(getFile('meta').filepath)
      const ga4        = parseGA4CSV(getFile('ga4').filepath)

      const { g, m } = calcTotals(googleRows, metaRows, ga4)

      const texts = await generateTexts({ g, m, ga4, marca, periodo, meta, mesAnt, contexto, proximos, histWhats, histTuo, histTel })

      const pptxBuffer = await buildPPTX({ g, m, ga4, marca, periodo, meta, mesAnt, histWhats, histTuo, histTel, texts })

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
      res.setHeader('Content-Disposition', `attachment; filename="Report_${marca}_${periodo}.pptx"`)
      res.send(pptxBuffer)

    } catch (e) {
      console.error(e)
      res.status(500).json({ message: e.message })
    }
  })
}
