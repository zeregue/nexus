'use client'

import { useState } from 'react'
import * as XLSX from 'xlsx'
import {
  AlertCircle,
  ArrowRight,
  BarChart3,
  Check,
  ChevronDown,
  CircleHelp,
  Download,
  FileSpreadsheet,
  FolderTree,
  LayoutDashboard,
  MoreHorizontal,
  MessageCircle,
  Package,
  Plus,
  Search,
  Settings2,
  ShieldCheck,
  Sparkles,
  Upload,
  X,
} from 'lucide-react'

type Product = { name: string; code: string; group: string; subgroup: string; ncm: string; cest: string; status: string }

const initialProducts: Product[] = [
  { name: 'Pão francês', code: '7891000100015', group: 'Padaria', subgroup: 'Pães', ncm: '1905.90.90', cest: '—', status: 'validado' },
  { name: 'Coca-Cola lata 350ml', code: '7894900011517', group: 'Bebidas', subgroup: 'Refrigerantes', ncm: '2202.10.00', cest: '03.007.00', status: 'validado' },
  { name: 'Coxinha de frango', code: '—', group: 'Salgados', subgroup: 'Fritos', ncm: '—', cest: '—', status: 'revisar' },
  { name: 'Brigadeiro tradicional', code: '—', group: 'Doces', subgroup: 'Doces de festa', ncm: '1704.90.90', cest: '—', status: 'revisar' },
  { name: 'Suco de laranja 300ml', code: '7896004400122', group: 'Bebidas', subgroup: 'Sucos', ncm: '2009.12.00', cest: '—', status: 'validado' },
]

const navItems = [
  { label: 'Visão geral', icon: LayoutDashboard },
  { label: 'Produtos', icon: Package, active: true, count: '248' },
  { label: 'Grupos e subgrupos', icon: FolderTree },
  { label: 'Dados fiscais', icon: ShieldCheck, count: '12' },
]

export default function Page() {
  const [activeTab, setActiveTab] = useState('Produtos')
  const [search, setSearch] = useState('')
  const [showUpload, setShowUpload] = useState(false)
  const [toast, setToast] = useState(false)
  const [selectedFile, setSelectedFile] = useState<string | null>(null)
  const [importKind, setImportKind] = useState<'estrutura' | 'produtos'>('estrutura')
  const [products, setProducts] = useState<Product[]>(initialProducts)
  const [importSummary, setImportSummary] = useState<string | null>(null)
  const [structureRows, setStructureRows] = useState<Array<{ group: string; count: number; subgroups: string[] }>>([])
  const [chatInput, setChatInput] = useState('')
  const [chatMessages, setChatMessages] = useState([{ role: 'assistant', text: 'Posso revisar a importação, localizar itens sem NCM ou sugerir correções na estrutura.' }])

  const visibleStructureRows = structureRows.length ? structureRows : Array.from(new Map(products.map((product) => [product.group, { group: product.group, count: products.filter((item) => item.group === product.group).length, subgroups: Array.from(new Set(products.filter((item) => item.group === product.group).map((item) => item.subgroup).filter((item) => item !== '—'))) }])).values())
  const filteredProducts = products.filter((product) =>
    product.name.toLowerCase().includes(search.toLowerCase()) || product.group.toLowerCase().includes(search.toLowerCase()),
  )

  function handleExport() {
    const header = ['Produto', 'Grupo', 'Subgrupo', 'NCM', 'EAN / GTIN', 'CEST', 'Status']
    const rows = filteredProducts.map((product) => [product.name, product.group, product.subgroup, product.ncm, product.code, product.cest, product.status])
    const csv = [header, ...rows].map((row) => row.map((cell) => `"${cell.replaceAll('"', '""')}"`).join(';')).join('\n')
    const blob = new Blob([`\ufeff${csv}`], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = 'produtos-organizados.csv'
    link.click()
    URL.revokeObjectURL(url)
    setToast(true)
    window.setTimeout(() => setToast(false), 2800)
  }

  const normalize = (value: unknown) => String(value ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  const firstValue = (row: Record<string, unknown>, names: string[]) => {
    const key = Object.keys(row).find((candidate) => names.includes(normalize(candidate)))
    return key ? String(row[key] ?? '').trim() : ''
  }

  function sendChatMessage() {
    const message = chatInput.trim()
    if (!message) return
    const lower = message.toLowerCase()
    const response = lower.includes('ncm') || lower.includes('cest')
      ? `${products.filter((product) => product.ncm === '—' || product.cest === '—').length} item(ns) ainda precisam de dados fiscais.`
      : lower.includes('grupo') || lower.includes('subgrupo')
        ? `Encontrei ${visibleStructureRows.length} grupo(s) na segunda aba. Os subgrupos foram promovidos para grupos.`
        : 'Posso ajudar a revisar grupos, produtos, códigos de barras, NCM e CEST/SEST.'
    setChatMessages((messages) => [...messages, { role: 'user', text: message }, { role: 'assistant', text: response }])
    setChatInput('')
  }

  async function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    setSelectedFile(file.name)
    const buffer = await file.arrayBuffer()
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
    const secondSheet = workbook.Sheets[workbook.SheetNames[1]]
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, { defval: '' })
    const secondRows = secondSheet ? XLSX.utils.sheet_to_json<Record<string, unknown>>(secondSheet, { defval: '' }) : []

    if (importKind === 'estrutura') {
      const organized = rows.map((row, index) => ({
        name: firstValue(row, ['item principal', 'grupo', 'grupo principal', 'categoria', 'departamento']) || `Item principal ${index + 1}`,
        code: firstValue(row, ['codigo', 'codigo do item']) || '—',
        group: 'Itens principais',
        subgroup: '—',
        ncm: '—',
        cest: '—',
        status: 'revisar',
      }))
      const promotedGroups = secondRows.map((row, index) => ({
        group: firstValue(row, ['subgrupo', 'grupos', 'grupo', 'categoria']) || `Grupo ${index + 1}`,
        count: 0,
        subgroups: [],
      }))
      setProducts(organized)
      setStructureRows(Array.from(new Map(promotedGroups.map((item) => [item.group, item])).values()))
      setImportSummary(`${organized.length} itens principais e ${promotedGroups.length} subgrupos convertidos em grupos na segunda aba.`)
    } else {
      const productSheet = workbook.Sheets['Itens principais'] || firstSheet
      const productRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(productSheet, { defval: '' })
      const merged = productRows.map((row) => ({
        name: firstValue(row, ['produto', 'descricao', 'descricao do produto', 'item']) || 'Produto sem descrição',
        code: firstValue(row, ['ean', 'gtin', 'codigo de barras', 'codigo do produto', 'codigo', 'cod barras']) || '—',
        group: firstValue(row, ['item principal', 'grupo', 'categoria', 'departamento']) || 'Sem grupo',
        subgroup: firstValue(row, ['subgrupo', 'subgrupo / categoria', 'familia', 'subcategoria']) || 'Sem subgrupo',
        ncm: firstValue(row, ['ncm', 'ncm fiscal']) || '—',
        cest: firstValue(row, ['cest', 'sest']) || '—',
        status: 'validado',
      }))
      setProducts(merged)
      setStructureRows([])
      setImportSummary(`${merged.length} produtos adicionados na ordem da planilha, com código de barras, NCM e CEST/SEST.`)
    }
    setToast(true)
    window.setTimeout(() => setToast(false), 2800)
  }

  return (
    <main className="min-h-screen bg-[#f7f8fa] text-[#18212f]">
      <aside className="fixed inset-y-0 left-0 hidden w-[244px] flex-col border-r border-[#e6e9ee] bg-white lg:flex">
        <div className="flex h-[76px] items-center gap-3 border-b border-[#eef0f3] px-6">
          <div className="flex h-9 w-9 items-center justify-center rounded-[10px] bg-[#1c6b5a] text-white shadow-sm"><Sparkles size={18} /></div>
          <div><div className="text-[15px] font-bold tracking-[-0.02em]">Organiza<span className="text-[#1c6b5a]">Planilha</span></div><div className="text-[10px] font-medium uppercase tracking-[0.14em] text-[#9aa3af]">Gestão de produtos</div></div>
        </div>
        <div className="px-3 pt-7">
          <div className="mb-3 px-3 text-[10px] font-bold uppercase tracking-[0.13em] text-[#a1a8b3]">Menu principal</div>
          <nav className="space-y-1">
            {navItems.map((item) => { const Icon = item.icon; return <button key={item.label} onClick={() => setActiveTab(item.label)} className={`flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left text-[13px] font-medium transition-colors ${activeTab === item.label ? 'bg-[#eaf5f1] text-[#17614f]' : 'text-[#687384] hover:bg-[#f5f7f8]'}`}><Icon size={17} strokeWidth={1.8} /><span className="flex-1">{item.label}</span>{item.count && <span className={`rounded-full px-2 py-0.5 text-[10px] ${activeTab === item.label ? 'bg-white text-[#17614f]' : 'bg-[#f0f2f4] text-[#8a93a0]'}`}>{item.count}</span>}</button> })}
          </nav>
          <div className="mb-3 mt-9 px-3 text-[10px] font-bold uppercase tracking-[0.13em] text-[#a1a8b3]">Atalhos</div>
          <button onClick={() => setShowUpload(true)} className="flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-[13px] font-medium text-[#687384] hover:bg-[#f5f7f8]"><Upload size={17} strokeWidth={1.8} />Importar planilha</button>
          <button onClick={handleExport} className="flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-[13px] font-medium text-[#687384] hover:bg-[#f5f7f8]"><Download size={17} strokeWidth={1.8} />Exportar resultado</button>
        </div>
        <div className="mt-auto border-t border-[#eef0f3] p-4"><div className="rounded-xl bg-[#f5faf8] p-3.5"><div className="mb-2 flex items-center gap-2 text-xs font-semibold text-[#235f50]"><CircleHelp size={15} />Precisa de ajuda?</div><p className="mb-3 text-[11px] leading-[1.5] text-[#7b9089]">Veja como preparar sua planilha para uma organização mais precisa.</p><button className="text-[11px] font-bold text-[#1c6b5a]">Acessar guia <ArrowRight className="ml-1 inline" size={12} /></button></div><div className="mt-4 flex items-center gap-3 px-2"><div className="flex h-8 w-8 items-center justify-center rounded-full bg-[#dcebe6] text-xs font-bold text-[#276b5a]">VC</div><div className="flex-1"><div className="text-xs font-semibold">Victor Castro</div><div className="text-[10px] text-[#9aa3af]">Administrador</div></div><Settings2 size={16} className="text-[#a1a8b3]" /></div></div>
      </aside>

      <section className="lg:pl-[244px]">
        <header className="flex h-[76px] items-center justify-between border-b border-[#e6e9ee] bg-white px-5 sm:px-8"><div className="flex items-center gap-3 lg:hidden"><div className="flex h-8 w-8 items-center justify-center rounded-lg bg-[#1c6b5a] text-white"><Sparkles size={15} /></div><span className="text-sm font-bold">OrganizaPlanilha</span></div><div className="hidden text-sm text-[#7d8794] lg:block">Quarta-feira, 12 de junho de 2024</div><div className="flex items-center gap-4"><button className="relative text-[#8c96a3]"><CircleHelp size={19} /><span className="absolute -right-1 -top-1 h-1.5 w-1.5 rounded-full bg-[#e78650]" /></button><div className="h-5 w-px bg-[#e5e8ec]" /><div className="flex items-center gap-2"><div className="flex h-8 w-8 items-center justify-center rounded-full bg-[#dcebe6] text-[11px] font-bold text-[#276b5a]">VC</div><span className="hidden text-xs font-semibold sm:block">Victor Castro</span><ChevronDown size={14} className="text-[#929aa5]" /></div></div></header>
        <div className="mx-auto max-w-[1360px] px-5 py-8 sm:px-8">
          <div className="mb-8 flex flex-col justify-between gap-4 sm:flex-row sm:items-end"><div><div className="mb-2 flex items-center gap-2 text-[11px] font-semibold text-[#1c6b5a]"><span>PROJETO ATUAL</span><span className="h-1 w-1 rounded-full bg-[#a9c7bd]" /><span className="text-[#9aa3af]">Loja do Vitor</span></div><h1 className="text-[27px] font-bold tracking-[-0.035em] text-[#172333]">Organização de produtos</h1><p className="mt-1 text-[13px] text-[#7f8997]">Revise e estruture seu cadastro antes de exportar.</p></div><div className="flex gap-2"><button onClick={() => setShowUpload(true)} className="inline-flex items-center justify-center gap-2 rounded-lg border border-[#dfe4e8] bg-white px-3.5 py-2.5 text-xs font-semibold text-[#506070] shadow-sm hover:bg-[#f8fafb]"><Upload size={15} />Importar nova</button><button onClick={handleExport} className="inline-flex items-center justify-center gap-2 rounded-lg bg-[#1c6b5a] px-3.5 py-2.5 text-xs font-semibold text-white shadow-sm shadow-[#1c6b5a]/20 hover:bg-[#155746]"><Download size={15} />Exportar planilha</button></div></div>
          <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4"><StatCard label="Total de produtos" value="248" detail="12 adicionados este mês" icon={<Package size={17} />} tone="green" /><StatCard label="Grupos principais" value="8" detail="2 grupos sem produtos" icon={<FolderTree size={17} />} tone="blue" /><StatCard label="Subgrupos cadastrados" value="24" detail="3 duplicidades encontradas" icon={<BarChart3 size={17} />} tone="orange" /><StatCard label="Precisam de revisão" value="12" detail="Dados fiscais incompletos" icon={<AlertCircle size={17} />} tone="red" /></div>
          <div className="mt-7 grid gap-5 xl:grid-cols-[1fr_310px]">
            <div className="rounded-xl border border-[#e3e7eb] bg-white shadow-[0_2px_8px_rgba(20,33,48,0.025)]"><div className="flex flex-col justify-between gap-4 border-b border-[#eef0f3] px-5 py-5 sm:flex-row sm:items-center"><div><h2 className="text-[15px] font-bold text-[#243142]">Produtos cadastrados</h2><p className="mt-1 text-xs text-[#929ba7]">Confira a classificação e os dados fiscais dos seus itens.</p></div><div className="relative"><Search className="absolute left-3 top-1/2 -translate-y-1/2 text-[#a8b0ba]" size={15} /><input aria-label="Buscar produtos" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar produto ou grupo" className="h-9 w-full rounded-lg border border-[#e1e5e9] bg-[#fbfcfd] pl-9 pr-3 text-xs outline-none transition focus:border-[#74aa9b] sm:w-[220px]" /></div></div><div className="overflow-x-auto"><table className="w-full min-w-[730px] text-left"><thead><tr className="border-b border-[#eef0f3] text-[10px] font-bold uppercase tracking-[0.08em] text-[#a0a8b3]"><th className="px-5 py-3 font-semibold">Produto</th><th className="py-3 font-semibold">Grupo / Subgrupo</th><th className="py-3 font-semibold">NCM</th><th className="py-3 font-semibold">EAN / GTIN</th><th className="py-3 font-semibold">Status</th><th className="px-5 py-3" /></tr></thead><tbody>{filteredProducts.map((product) => <tr key={product.name} className="border-b border-[#f0f2f4] last:border-0 hover:bg-[#fcfdfd]"><td className="px-5 py-3.5"><div className="text-xs font-semibold text-[#2c3847]">{product.name}</div><div className="mt-1 text-[10px] text-[#a0a8b3]">CEST: {product.cest}</div></td><td><div className="text-xs text-[#4c5968]">{product.group}</div><div className="mt-1 text-[10px] text-[#a0a8b3]">{product.subgroup}</div></td><td className="font-mono text-[11px] text-[#596675]">{product.ncm}</td><td className="font-mono text-[11px] text-[#7a8591]">{product.code}</td><td>{product.status === 'validado' ? <span className="inline-flex items-center gap-1 rounded-full bg-[#eaf6f0] px-2 py-1 text-[10px] font-semibold text-[#237153]"><Check size={11} />Validado</span> : <span className="inline-flex items-center gap-1 rounded-full bg-[#fff4e9] px-2 py-1 text-[10px] font-semibold text-[#b56b28]"><AlertCircle size={11} />Revisar</span>}</td><td className="px-5 text-right"><button aria-label={`Mais opções para ${product.name}`} className="text-[#a3abb5] hover:text-[#516172]"><MoreHorizontal size={17} /></button></td></tr>)}</tbody></table>{filteredProducts.length === 0 && <div className="px-5 py-10 text-center text-sm text-[#8a94a1]">Nenhum produto encontrado.</div>}</div><div className="flex items-center justify-between border-t border-[#eef0f3] px-5 py-3"><span className="text-[11px] text-[#9aa3ae]">Exibindo {filteredProducts.length} de 248 produtos</span><button className="text-[11px] font-semibold text-[#1c6b5a]">Ver todos <ArrowRight className="ml-1 inline" size={12} /></button></div></div>
            <div className="space-y-5"><div className="rounded-xl border border-[#e3e7eb] bg-white p-5 shadow-[0_2px_8px_rgba(20,33,48,0.025)]"><div className="mb-4 flex items-center justify-between"><div><h2 className="text-[14px] font-bold text-[#243142]">Estrutura atual</h2><p className="mt-1 text-[11px] text-[#929ba7]">Distribuição por grupo principal</p></div><FolderTree size={17} className="text-[#75a497]" /></div><div className="space-y-3">{[['Alimentos', 76, 'bg-[#5f9f8a]'], ['Bebidas', 58, 'bg-[#82b5a6]'], ['Doces', 44, 'bg-[#b49a72]'], ['Salgados', 38, 'bg-[#d2ad75]'], ['Padaria', 32, 'bg-[#d9c49f]']].map(([name, value, color]) => <div key={String(name)}><div className="mb-1.5 flex justify-between text-[11px]"><span className="font-medium text-[#596675]">{name}</span><span className="text-[#939ca7]">{value}</span></div><div className="h-1.5 rounded-full bg-[#edf0f1]"><div className={`h-1.5 rounded-full ${color}`} style={{ width: `${(Number(value) / 76) * 100}%` }} /></div></div>)}</div><button className="mt-5 w-full rounded-lg border border-[#e3e8e7] py-2 text-[11px] font-semibold text-[#377563] hover:bg-[#f4faf7]">Gerenciar grupos <ArrowRight className="ml-1 inline" size={12} /></button></div><div className="rounded-xl border border-[#f3dfc5] bg-[#fffbf5] p-5"><div className="mb-3 flex items-center gap-2 text-[12px] font-bold text-[#95612b]"><AlertCircle size={16} />Atenção necessária</div><p className="text-[11px] leading-[1.55] text-[#997959]">12 produtos estão sem NCM ou CEST. Revise os itens antes de exportar para evitar inconsistências fiscais.</p><button className="mt-3 text-[11px] font-bold text-[#b56b28]">Revisar agora <ArrowRight className="ml-1 inline" size={12} /></button></div></div>
          </div>
        </div>
        <section className="mx-auto max-w-[1360px] px-5 pb-8 sm:px-8">
          <div className="rounded-xl border border-[#e3e7eb] bg-white shadow-[0_2px_8px_rgba(20,33,48,0.025)]">
            <div className="flex items-center justify-between border-b border-[#eef0f3] px-5 py-4"><div><h2 className="text-[15px] font-bold text-[#243142]">Grupos e subgrupos</h2><p className="mt-1 text-xs text-[#929ba7]">Segunda aba: os subgrupos importados foram promovidos a grupos.</p></div><FolderTree size={17} className="text-[#75a497]" /></div>
            <div className="overflow-x-auto"><table className="w-full min-w-[560px] text-left"><thead><tr className="border-b border-[#eef0f3] text-[10px] font-bold uppercase tracking-[0.08em] text-[#a0a8b3]"><th className="px-5 py-3">Grupo</th><th className="py-3">Itens</th><th className="py-3">Subgrupos encontrados</th></tr></thead><tbody>{visibleStructureRows.map((row) => <tr key={row.group} className="border-b border-[#f0f2f4] last:border-0"><td className="px-5 py-3 text-xs font-semibold text-[#2c3847]">{row.group}</td><td className="py-3 text-xs text-[#596675]">{row.count}</td><td className="py-3 text-xs text-[#596675]">{row.subgroups.length ? row.subgroups.join(', ') : 'Sem subgrupo'}</td></tr>)}</tbody></table></div>
          </div>
          <div className="mt-5 rounded-xl border border-[#dce9e4] bg-[#f8fcfa] p-5"><div className="mb-3 flex items-center gap-2 text-sm font-bold text-[#245d4e]"><MessageCircle size={17} />Assistente da planilha</div><div className="max-h-44 space-y-2 overflow-y-auto rounded-lg bg-white p-3">{chatMessages.map((message, index) => <div key={`${message.role}-${index}`} className={`max-w-[88%] rounded-lg px-3 py-2 text-xs ${message.role === 'user' ? 'ml-auto bg-[#1c6b5a] text-white' : 'bg-[#edf6f2] text-[#3e6156]'}`}>{message.text}</div>)}</div><div className="mt-3 flex gap-2"><input aria-label="Mensagem para o assistente" value={chatInput} onChange={(event) => setChatInput(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter' && !event.nativeEvent.isComposing && event.keyCode !== 229) sendChatMessage() }} placeholder="Pergunte sobre os itens importados..." className="h-10 flex-1 rounded-lg border border-[#dce5e1] bg-white px-3 text-xs outline-none focus:border-[#74aa9b]" /><button aria-label="Enviar mensagem" onClick={sendChatMessage} className="flex h-10 w-10 items-center justify-center rounded-lg bg-[#1c6b5a] text-white hover:bg-[#155746]"><ArrowRight size={16} /></button></div></div>
        </section>
      </section>
      {showUpload && <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#13251f]/35 px-4"><div className="w-full max-w-[470px] rounded-2xl bg-white p-6 shadow-2xl"><div className="mb-5 flex items-start justify-between"><div><h2 className="text-lg font-bold text-[#243142]">Importar planilha</h2><p className="mt-1 text-xs text-[#87919d]">Escolha a etapa e envie sua planilha Excel.</p></div><button aria-label="Fechar" onClick={() => setShowUpload(false)} className="text-[#9aa3ae] hover:text-[#435160]"><X size={18} /></button></div><div className="mb-4 grid grid-cols-2 gap-2"><button onClick={() => setImportKind('estrutura')} className={`rounded-lg border px-3 py-2 text-left text-[11px] font-semibold ${importKind === 'estrutura' ? 'border-[#8fc1b0] bg-[#edf8f3] text-[#236b57]' : 'border-[#e3e7eb] text-[#7f8997]'}`}>1. Estrutura<span className="mt-1 block text-[10px] font-normal opacity-75">Grupos e subgrupos</span></button><button onClick={() => setImportKind('produtos')} className={`rounded-lg border px-3 py-2 text-left text-[11px] font-semibold ${importKind === 'produtos' ? 'border-[#8fc1b0] bg-[#edf8f3] text-[#236b57]' : 'border-[#e3e7eb] text-[#7f8997]'}`}>2. Produtos<span className="mt-1 block text-[10px] font-normal opacity-75">EAN, NCM e CEST</span></button></div><div className="rounded-xl border-2 border-dashed border-[#b9d6cc] bg-[#f7fcfa] px-6 py-9 text-center"><div className="mx-auto mb-3 flex h-11 w-11 items-center justify-center rounded-full bg-[#e4f3ed] text-[#2f806b]"><FileSpreadsheet size={22} /></div><p className="text-sm font-semibold text-[#385348]">Arraste seu arquivo aqui</p><p className="mt-1 text-xs text-[#96a59f]">ou selecione um arquivo .xlsx ou .xls</p><label className="mt-4 inline-flex cursor-pointer rounded-lg bg-[#1c6b5a] px-4 py-2 text-xs font-semibold text-white hover:bg-[#155746]">Selecionar arquivo<input className="sr-only" type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} /></label>{selectedFile && <p className="mt-3 text-xs font-medium text-[#2f806b]">Arquivo selecionado: {selectedFile}</p>}</div><div className="mt-5 flex items-start gap-2 text-[11px] leading-[1.5] text-[#86908e]"><ShieldCheck size={14} className="mt-0.5 shrink-0 text-[#5f9c8a]" />Seus dados são processados com segurança e os códigos fiscais originais nunca são inventados.</div></div></div>}
      {toast && <div role="status" className="fixed bottom-6 right-6 z-50 flex items-center gap-2 rounded-lg bg-[#183d33] px-4 py-3 text-xs font-semibold text-white shadow-xl"><Check size={15} />Exportação preparada com sucesso</div>}
    </main>
  )
}

function StatCard({ label, value, detail, icon, tone }: { label: string; value: string; detail: string; icon: React.ReactNode; tone: string }) {
  const tones: Record<string, string> = { green: 'bg-[#e7f4ef] text-[#2c806c]', blue: 'bg-[#eaf1f8] text-[#4b7ca5]', orange: 'bg-[#faf1e6] text-[#b47b3d]', red: 'bg-[#fcedeb] text-[#bb665d]' }
  return <div className="rounded-xl border border-[#e3e7eb] bg-white p-4 shadow-[0_2px_8px_rgba(20,33,48,0.025)]"><div className="flex items-start justify-between"><div className="text-[11px] font-semibold text-[#89939f]">{label}</div><div className={`flex h-7 w-7 items-center justify-center rounded-lg ${tones[tone]}`}>{icon}</div></div><div className="mt-3 text-[25px] font-bold tracking-[-0.04em] text-[#263343]">{value}</div><div className="mt-1 text-[10px] text-[#9aa4af]">{detail}</div></div>
}
