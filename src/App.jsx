import React, { useEffect, useMemo, useState } from 'react'
import { WalletCards, Landmark, PiggyBank, ReceiptText, Gauge, TrendingUp, Plus, Pencil, Trash2, Search, RotateCcw, Upload, FileSpreadsheet, X, Download, DatabaseBackup } from 'lucide-react'
import * as XLSX from 'xlsx'

const INITIAL = [
  {id:1,date:'2026-07-01',description:'회사 월급',amount:3495434,type:'Income',category:'급여'},
  {id:2,date:'2026-07-01',description:'양주 월세 지원',amount:200000,type:'Income',category:'기타수입'},
  {id:3,date:'2026-07-01',description:'국민 청년드림적금',amount:500000,type:'Saving',category:'저축'},
  {id:4,date:'2026-07-01',description:'우리 주택청약',amount:100000,type:'Saving',category:'저축'},
  {id:5,date:'2026-07-01',description:'KB연금저축',amount:200000,type:'Saving',category:'저축'},
  {id:6,date:'2026-07-01',description:'삼성 ISA',amount:200000,type:'Saving',category:'저축'},
  {id:7,date:'2026-07-01',description:'월세',amount:415000,type:'Fixed',category:'주거'},
  {id:8,date:'2026-07-01',description:'농협종합보험',amount:83671,type:'Fixed',category:'보험'},
  {id:9,date:'2026-07-01',description:'리아나 치과보험',amount:36500,type:'Fixed',category:'보험'},
  {id:10,date:'2026-07-01',description:'푸본현대보험',amount:34160,type:'Fixed',category:'보험'},
  {id:11,date:'2026-07-01',description:'SK 인터넷',amount:17600,type:'Fixed',category:'통신'},
  {id:12,date:'2026-07-01',description:'SKT 통신요금',amount:38670,type:'Fixed',category:'통신'},
  {id:13,date:'2026-07-01',description:'유튜브 프리미엄',amount:14900,type:'Fixed',category:'구독'},
  {id:14,date:'2026-07-01',description:'웨이브',amount:10900,type:'Fixed',category:'구독'},
  {id:15,date:'2026-07-01',description:'LG 렌탈비',amount:20900,type:'Fixed',category:'렌탈'},
  {id:16,date:'2026-07-01',description:'쿠팡 와우',amount:7890,type:'Fixed',category:'구독'},
  {id:17,date:'2026-07-01',description:'와이즐리 구독',amount:2990,type:'Fixed',category:'구독'},
  {id:18,date:'2026-07-01',description:'유니세프',amount:20000,type:'Fixed',category:'기부'},
  {id:19,date:'2026-07-01',description:'레고 게임 구독',amount:24000,type:'Fixed',category:'구독'},
  {id:20,date:'2026-07-01',description:'교통통장',amount:50000,type:'Fixed',category:'교통'},
  {id:21,date:'2026-07-01',description:'마이너스통장 이자',amount:90000,type:'Fixed',category:'대출이자'},
  {id:22,date:'2026-07-01',description:'용산운전센터 운동',amount:105000,type:'Fixed',category:'운동'},
  {id:23,date:'2026-07-01',description:'우피아 모임',amount:10000,type:'Fixed',category:'모임'},
  {id:24,date:'2026-07-01',description:'당뇨 치료 병원',amount:35000,type:'Fixed',category:'병원'}
]

const TYPE_LABELS={Income:'수입',Saving:'저축',Fixed:'고정비',Variable:'변동비'}
const CATEGORY_MAP={
  Income:['급여','기타수입'],
  Saving:['ISA','적금','청약','투자','기타저축'],
  Fixed:['주거','보험','운동','대출이자','구독','통신','교통','렌탈','기부','모임','병원','기타'],
  Variable:['식비','카페','쇼핑','생활용품','공과금','취미','교통','의료','자기계발','사업','기타']
}
const VAR_CATS=CATEGORY_MAP.Variable
const ALL_CATS=[...new Set(Object.values(CATEGORY_MAP).flat())]
const MANUAL_CATEGORIES=CATEGORY_MAP
const ASSETS=47983679
const DEBTS=19195089
const APP_VERSION='0.10.2'

const won=n=>new Intl.NumberFormat('ko-KR',{style:'currency',currency:'KRW',maximumFractionDigits:0}).format(Number(n)||0)
const monthKey=date=>(date||'').slice(0,7)
const latestMonth=items=>{
  const months=[...new Set((items||[]).map(t=>monthKey(t.date)).filter(Boolean))]
  return months.sort().reverse()[0]||'2026-07'
}

function Stat({label,value,sub,Icon}){
  return <div className="stat"><div className="statTop"><span>{label}</span><Icon size={18}/></div><strong>{value}</strong>{sub&&<small>{sub}</small>}</div>
}

export default function App(){
  const [tab,setTab]=useState('dashboard')
  const [transactions,setTransactions]=useState(()=>{
    try{
      const v3=localStorage.getItem('wonseok-finance-v03')
      if(v3) return JSON.parse(v3)
      const old=localStorage.getItem('wonseok-finance-v021')
      return old?JSON.parse(old):INITIAL
    }catch{return INITIAL}
  })
  const [ledger,setLedger]=useState(()=>{
    try{
      const saved=localStorage.getItem('wonseok-finance-ledger-v1')
      if(saved)return JSON.parse(saved)
      // 기존 버전 사용자는 현재 거래 데이터를 최초 확정 원장으로 자동 이관
      const legacy=localStorage.getItem('wonseok-finance-v03')
      if(legacy)return JSON.parse(legacy)
      const old=localStorage.getItem('wonseok-finance-v021')
      return old?JSON.parse(old):INITIAL
    }catch{return INITIAL}
  })
  const [selectedMonth,setSelectedMonth]=useState(()=>latestMonth(ledger))
  const [editingId,setEditingId]=useState(null)
  const [form,setForm]=useState({date:'2026-08-11',description:'',amount:'',type:'Variable',category:'식비'})
  const [search,setSearch]=useState('')
  const [typeFilter,setTypeFilter]=useState('All')
  const [monthFilter,setMonthFilter]=useState('All')
  const [importPreview,setImportPreview]=useState([])
  const [importName,setImportName]=useState('')
  const [importError,setImportError]=useState('')
  const [importStats,setImportStats]=useState(null)
  const [categoryRules,setCategoryRules]=useState(()=>{
    try{return JSON.parse(localStorage.getItem('wonseok-finance-category-rules')||'{}')}catch{return {}}
  })
  const [installments,setInstallments]=useState(()=>{
    try{return JSON.parse(localStorage.getItem('wonseok-finance-installments')||'[]')}catch{return []}
  })
  const [installmentForm,setInstallmentForm]=useState({description:'',card:'현대카드',totalAmount:'',months:'',startMonth:'2026-08'})
  const [installmentCandidates,setInstallmentCandidates]=useState([])
  const [duplicateOverrides,setDuplicateOverrides]=useState({})
  const [backupError,setBackupError]=useState('')

  const save=next=>{
    setTransactions(next)
    localStorage.setItem('wonseok-finance-v03',JSON.stringify(next))
  }
  const saveLedger=next=>{
    setLedger(next)
    localStorage.setItem('wonseok-finance-ledger-v1',JSON.stringify(next))
  }

  const availableMonths=useMemo(()=>{
    const months=[...new Set(ledger.map(t=>monthKey(t.date)).filter(Boolean))]
    return months.sort().reverse()
  },[ledger])

  useEffect(()=>{
    if(!availableMonths.length)return
    if(!availableMonths.includes(selectedMonth)){
      setSelectedMonth(availableMonths[0])
    }
  },[availableMonths,selectedMonth])

  const monthTransactions=useMemo(
    ()=>ledger.filter(t=>monthKey(t.date)===selectedMonth),
    [ledger,selectedMonth]
  )

  const totals=useMemo(()=>{
    const sum=type=>monthTransactions.filter(t=>t.type===type).reduce((a,b)=>a+Number(b.amount||0),0)
    const income=sum('Income'), saving=sum('Saving'), fixed=sum('Fixed'), variable=sum('Variable')
    const budget=income-saving-fixed
    return {income,saving,fixed,variable,budget,remaining:budget-variable,net:ASSETS-DEBTS}
  },[monthTransactions])

  const fixedBy=useMemo(()=>{
    const m={}
    monthTransactions.filter(t=>t.type==='Fixed').forEach(t=>m[t.category]=(m[t.category]||0)+Number(t.amount||0))
    return Object.entries(m).sort((a,b)=>b[1]-a[1])
  },[monthTransactions])

  const variableBy=useMemo(()=>VAR_CATS.map(cat=>[
    cat,
    monthTransactions.filter(t=>t.type==='Variable'&&t.category===cat).reduce((s,t)=>s+Number(t.amount||0),0)
  ]),[monthTransactions])

  const importMonths=useMemo(()=>{
    const months=[...new Set(transactions.map(t=>monthKey(t.date)).filter(Boolean))]
    return months.sort().reverse()
  },[transactions])

  const filteredTransactions=useMemo(()=>{
    const q=search.trim().toLowerCase()
    return transactions
      .filter(t=>monthFilter==='All'||monthKey(t.date)===monthFilter)
      .filter(t=>typeFilter==='All'||t.type===typeFilter)
      .filter(t=>!q||t.description.toLowerCase().includes(q)||t.category.toLowerCase().includes(q))
      .sort((a,b)=>b.date.localeCompare(a.date)||Number(b.id)-Number(a.id))
  },[transactions,search,typeFilter,monthFilter])

  const submit=e=>{
    e.preventDefault()
    if(!form.description.trim()||!Number(form.amount)) return
    const item={...form,amount:Number(form.amount),paymentMethod:'수동입력',source:'수동입력',needsReview:false}
    if(editingId){
      save(transactions.map(t=>t.id===editingId?{...item,id:editingId}:t))
      setEditingId(null)
    }else{
      save([...transactions,{...item,id:Date.now()}])
    }
    setForm({...form,description:'',amount:''})
  }

  const edit=t=>{
    setForm({date:t.date,description:t.description,amount:t.amount,type:t.type,category:t.category})
    setEditingId(t.id)
    setTab('transactions')
  }

  const remove=id=>{
    if(confirm('이 거래를 삭제할까요?')) save(transactions.filter(t=>t.id!==id))
  }

  const resetFilters=()=>{
    setSearch('')
    setTypeFilter('All')
    setMonthFilter('All')
  }

  const cellText=v=>v==null?'':String(v).trim()
  const parseAmount=v=>{
    if(typeof v==='number') return Math.abs(v)
    const n=Number(cellText(v).replace(/[^0-9.-]/g,''))
    return Number.isFinite(n)?Math.abs(n):0
  }
  const excelDate=v=>{
    if(v instanceof Date) return v.toISOString().slice(0,10)
    if(typeof v==='number'){
      const d=XLSX.SSF.parse_date_code(v)
      if(d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`
    }
    const t=cellText(v)
      .replace(/\s+/g,'')
      .replace(/년|월/g,'-')
      .replace(/일/g,'')
      .replace(/[./]/g,'-')
    const m=t.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/)
    return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:''
  }

  const classifyHyundai=(merchant,amount)=>{
    const m=merchant.toUpperCase()
    if(m.includes('SKT')) return {type:'Fixed',category:'통신',review:false}
    if(m.includes('GS25')) return {type:'Variable',category:'생활용품',review:false}
    if(m.includes('써브웨이')||m.includes('김밥')) return {type:'Variable',category:'식비',review:false}
    if(m.includes('커피')) return {type:'Variable',category:'카페',review:false}
    if(m.includes('그린카')||m.includes('오일뱅크')||m.includes('주유소')) return {type:'Variable',category:'교통',review:false}
    if(m.includes('무신사')) return {type:'Variable',category:'쇼핑',review:false}
    if(m.includes('STEAM')||m.includes('XSOLLA')||m.includes('EPICGAMES')) return {type:'Variable',category:'취미',review:false}
    if(m.includes('웨이브')) return {type:'Fixed',category:'구독',review:true}
    if(m.includes('쿠팡') && amount===7890) return {type:'Fixed',category:'구독',review:true}
    return {type:'Variable',category:'기타',review:true}
  }

  const classifyKB=(merchant,amount)=>{
    const m=merchant.toUpperCase()
    if(m.includes('SK브로드밴드')||m.includes('SKT')) return {type:'Fixed',category:'통신',review:false}
    if(m.includes('GS25')) return {type:'Variable',category:'생활용품',review:false}
    if(m.includes('파리바게뜨')||m.includes('커피')||m.includes('남산왕돈까스')) return {type:'Variable',category:'식비',review:false}
    if(m.includes('카카오T')||m.includes('자동결제')) return {type:'Variable',category:'교통',review:true}
    if(m.includes('LG헬로비전')) return {type:'Fixed',category:'구독',review:true}
    if(m.includes('쿠팡')) return {type:'Variable',category:'쇼핑',review:true}
    return {type:'Variable',category:'기타',review:true}
  }

  const parseHyundaiHtml=text=>{
    const doc=new DOMParser().parseFromString(text,'text/html')
    const tableRows=[...doc.querySelectorAll('tr')]
    let headerIndex=-1,headers=[]
    tableRows.forEach((tr,i)=>{
      const cells=[...tr.querySelectorAll('th,td')].map(c=>cellText(c.textContent))
      if(cells.includes('이용일')&&cells.includes('이용가맹점')&&cells.includes('이용금액')){headerIndex=i;headers=cells}
    })
    if(headerIndex<0)return null
    const idx=name=>headers.indexOf(name)
    const dateIdx=idx('이용일'),merchantIdx=idx('이용가맹점'),amountIdx=idx('이용금액'),installmentIdx=idx('할부/회차'),principalIdx=idx('결제원금')
    const rows=[],installmentRows=[]
    let skippedInstallmentPayments=0,skippedSummaryRows=0
    tableRows.slice(headerIndex+1).forEach((tr,i)=>{
      const cells=[...tr.querySelectorAll('td,th')].map(c=>cellText(c.textContent))
      if(!cells.length)return
      const date=excelDate(cells[dateIdx]),merchant=cellText(cells[merchantIdx])
      if(!date||!merchant){if(merchant.includes('소계')||merchant.includes('합계'))skippedSummaryRows++;return}
      const originalAmount=parseAmount(cells[amountIdx]),billedPrincipal=parseAmount(cells[principalIdx]),installment=cellText(cells[installmentIdx])
      let amount=originalAmount||billedPrincipal,installmentInfo=null
      if(installment){
        const mt=installment.match(/(\d+)\s*\/\s*(\d+)/)
        if(mt){
          const total=Number(mt[1]),round=Number(mt[2])
          installmentInfo={total,round,billedPrincipal,originalAmount}
          installmentRows.push({key:`현대카드|${merchant}|${originalAmount}|${total}`,description:merchant,card:'현대카드',totalAmount:originalAmount,months:total,currentRound:round,billedPrincipal,statementDate:date,startMonth:''})
          if(round!==1){skippedInstallmentPayments++;return}
          amount=originalAmount
        }
      }
      if(!amount)return
      const cls=classifyHyundai(merchant,amount)
      rows.push({id:`preview-hyundai-${i}`,date,description:merchant,amount,type:cls.type,category:cls.category,paymentMethod:'현대카드',source:'현대카드 명세서',needsReview:cls.review,installment:installmentInfo})
    })
    return {rows,installmentRows,stats:{parser:'현대카드 전용',imported:rows.length,skippedInstallmentPayments,skippedSummaryRows}}
  }

  const parseKBWorkbook=wb=>{
    let best=null
    wb.SheetNames.forEach(name=>{
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true})
      let headerIndex=-1,headers=[]
      rows.forEach((r,i)=>{
        const cells=r.map(cellText)
        if(
          cells.includes('이용일') &&
          cells.includes('이용하신곳') &&
          cells.some(x=>x.includes('국내이용금액'))
        ){
          headerIndex=i
          headers=cells
        }
      })
      if(headerIndex<0)return

      const dateIdx=headers.indexOf('이용일')
      const merchantIdx=headers.indexOf('이용하신곳')
      const amountIdx=headers.findIndex(x=>x.includes('국내이용금액'))
      const statusIdx=headers.indexOf('상태')

      const parsed=[]
      let skippedCanceled=0

      rows.slice(headerIndex+1).forEach((r,i)=>{
        const date=excelDate(r[dateIdx])
        const merchant=cellText(r[merchantIdx])
        const amount=parseAmount(r[amountIdx])
        const status=statusIdx>=0 ? cellText(r[statusIdx]) : ''

        if(!date || !merchant || !amount)return

        // 국민카드 원본의 취소 상태: 승인취소, 취소전표매입 등
        if(status.includes('취소')){
          skippedCanceled++
          return
        }

        const cls=classifyKB(merchant,amount)
        parsed.push({
          id:`preview-kb-${i}`,
          date,
          description:merchant,
          amount,
          type:cls.type,
          category:cls.category,
          paymentMethod:'국민카드',
          source:'국민카드 이용내역',
          needsReview:cls.review
        })
      })

      if(!best || parsed.length>best.rows.length){
        best={
          rows:parsed,
          stats:{
            parser:'국민카드 전용',
            imported:parsed.length,
            skippedInstallmentPayments:0,
            skippedSummaryRows:skippedCanceled
          }
        }
      }
    })
    return best
  }

  const classifyToss=(description,txType,amount)=>{
    const m=description.toUpperCase()

    if(txType==='대출이자') return {type:'Fixed',category:'대출이자',review:false}
    if(m.includes('유니세프')) return {type:'Fixed',category:'기부',review:false}
    if(m.includes('GS25')||m.includes('지에스25')||m.includes('씨유')||m.includes('세븐일레븐')) return {type:'Variable',category:'생활용품',review:false}
    if(m.includes('한솥')||m.includes('버거리')||m.includes('소금빵')||m.includes('베이크')||m.includes('맥코이')||m.includes('카레')) return {type:'Variable',category:'식비',review:false}
    if(m.includes('약국')||m.includes('내과')||m.includes('의원')||m.includes('병원')) return {type:'Variable',category:'의료',review:false}
    if(m.includes('PC방')) return {type:'Variable',category:'취미',review:false}

    return {type:txType==='자동이체'?'Fixed':'Variable',category:'기타',review:true}
  }

  const isTossInternalTransfer=(desc,txType,institution)=>{
    const d=desc.toUpperCase()
    const inst=institution.toUpperCase()

    // 신용카드 대금은 이미 카드 개별 사용 내역에서 소비로 기록
    if(d.includes('현대카드')||d.includes('국민카드')||d.includes('KB카드')) return {exclude:true,reason:'카드대금'}

    // 투자/증권 계좌로의 자금 이동은 소비가 아님
    if(d.includes('토스증권')||inst.includes('토스증권')) return {exclude:true,reason:'투자계좌 이동'}

    // 전자지갑 충전은 실제 사용처에서 소비를 잡기 위해 제외
    if(d.includes('토스페이충전')) return {exclude:true,reason:'전자지갑 충전'}

    // ATM 출금은 현금으로 형태만 바뀐 것이므로 소비에서 제외
    if(txType==='ATM출금') return {exclude:true,reason:'현금 인출'}

    // 본인 명의 다른 계좌로의 이체
    if(d==='정원석' && ['케이뱅크','KB국민은행','카카오뱅크','우리은행'].some(x=>inst.includes(x.toUpperCase()))) {
      return {exclude:true,reason:'내 계좌 이동'}
    }

    // 공동통장 적립도 소비가 아니라 자금 이동
    if(d.includes('공동 통장')) return {exclude:true,reason:'공동계좌 이동'}

    // 쿠팡/토스뱅크 펌뱅킹 10만원 단위는 페이머니 충전 성격으로 보고 제외
    if(txType==='펌뱅킹출금' && inst.includes('토스뱅크') && (d==='쿠팡'||d.includes('토스페이'))) {
      return {exclude:true,reason:'페이머니 충전'}
    }

    return {exclude:false,reason:''}
  }

  const parseTossWorkbook=wb=>{
    let best=null

    wb.SheetNames.forEach(name=>{
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true})
      let headerIndex=-1,headers=[]

      rows.forEach((r,i)=>{
        const cells=r.map(cellText)
        if(cells.includes('거래 일시')&&cells.includes('적요')&&cells.includes('거래 유형')&&cells.includes('거래 금액')){
          headerIndex=i
          headers=cells
        }
      })
      if(headerIndex<0)return

      const dateIdx=headers.indexOf('거래 일시')
      const descIdx=headers.indexOf('적요')
      const typeIdx=headers.indexOf('거래 유형')
      const institutionIdx=headers.indexOf('거래 기관')
      const amountIdx=headers.indexOf('거래 금액')

      const parsed=[]
      const excluded=[]
      let incomeSkipped=0

      rows.slice(headerIndex+1).forEach((r,i)=>{
        const rawDate=cellText(r[dateIdx])
        const date=excelDate(rawDate)
        const description=cellText(r[descIdx])
        const txType=cellText(r[typeIdx])
        const institution=institutionIdx>=0?cellText(r[institutionIdx]):''
        const signedAmount=Number(String(r[amountIdx]??'').replace(/[^0-9.-]/g,''))
        const amount=Math.abs(signedAmount||0)

        if(!date||!description||!amount)return

        // 입금은 소비 업로드에서 제외. 수입은 기존 수입 모듈에서 별도로 관리.
        if(signedAmount>0 || txType.includes('입금')){
          incomeSkipped++
          return
        }

        const internal=isTossInternalTransfer(description,txType,institution)
        if(internal.exclude){
          excluded.push({description,amount,reason:internal.reason})
          return
        }

        let paymentMethod='계좌이체'
        if(txType==='체크카드결제') paymentMethod='토스체크'

        const cls=classifyToss(description,txType,amount)

        // 사용자 데이터 기준 월세 자동이체는 식별 가능하므로 1차 자동분류
        let type=cls.type,category=cls.category,review=cls.review
        if(txType==='자동이체' && description==='김인수' && amount===415000){
          type='Fixed';category='주거';review=false
        }

        parsed.push({
          id:`preview-toss-${i}`,
          date,
          description,
          amount,
          type,
          category,
          paymentMethod,
          source:'토스뱅크 거래내역',
          needsReview:review,
          transactionType:txType,
          institution
        })
      })

      if(!best||parsed.length>best.rows.length){
        best={
          rows:parsed,
          stats:{
            parser:'토스 전용',
            imported:parsed.length,
            skippedInstallmentPayments:0,
            skippedSummaryRows:excluded.length,
            excludedTransfers:excluded.length,
            incomeSkipped,
            excluded
          }
        }
      }
    })
    return best
  }

  const guessRows=rows=>{
    const out=[]
    rows.forEach((r,i)=>{
      const vals=r.map(cellText)
      const dateIndex=r.findIndex(v=>excelDate(v))
      if(dateIndex<0)return
      const date=excelDate(r[dateIndex])
      let amount=0
      for(let j=r.length-1;j>=0;j--){
        const a=parseAmount(r[j])
        if(a>=100&&!excelDate(r[j])){amount=a;break}
      }
      if(!amount)return
      const candidates=vals.filter((v,j)=>j!==dateIndex&&v&&!/^[-+]?\d[\d,.]*$/.test(v)&&!/^(일시불|국내|해외|정상|승인|취소)/.test(v))
      const description=candidates.sort((a,b)=>b.length-a.length)[0]||'업로드 거래'
      out.push({id:`preview-${i}`,date,description,amount,type:'Variable',category:'기타',paymentMethod:'미지정',source:'파일 업로드',needsReview:true})
    })
    const seen=new Set()
    return out.filter(x=>{const k=`${x.date}|${x.description}|${x.amount}`;if(seen.has(k))return false;seen.add(k);return true}).slice(0,300)
  }

  const handleFile=async e=>{
    const file=e.target.files?.[0]
    if(!file)return
    setImportName(file.name);setImportError('');setImportPreview([]);setImportStats(null);setInstallmentCandidates([])
    try{
      const text=await file.text()
      const looksHyundai=text.includes('이용가맹점')&&text.includes('이용금액')&&text.includes('현대카드')
      if(looksHyundai){
        const parsed=parseHyundaiHtml(text)
        if(!parsed||!parsed.rows.length)throw new Error('현대카드 거래 행을 찾지 못했습니다.')
        setImportPreview(parsed.rows)
        setImportStats(parsed.stats)
      }else{
        const data=await file.arrayBuffer()
        const wb=XLSX.read(data,{type:'array',cellDates:true})
        const kb=parseKBWorkbook(wb)
        const toss=parseTossWorkbook(wb)
        if(kb&&kb.rows.length){
          setImportPreview(applyCategoryRules(kb.rows))
          setImportStats(kb.stats)
        }else if(toss&&toss.rows.length){
          setImportPreview(applyCategoryRules(toss.rows))
          setImportStats(toss.stats)
        }else{
          let best=[]
          wb.SheetNames.forEach(name=>{
            const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true})
            const parsed=guessRows(rows)
            if(parsed.length>best.length)best=parsed
          })
          if(!best.length)throw new Error('거래 행을 찾지 못했습니다.')
          setImportPreview(applyCategoryRules(best))
          setImportStats({parser:'범용',imported:best.length,skippedInstallmentPayments:0,skippedSummaryRows:0})
        }
      }
    }catch(err){setImportError(`파일을 읽지 못했습니다: ${err.message}`)}
    e.target.value=''
  }

  const norm=v=>(v||'').trim().toUpperCase().replace(/\s+/g,' ')
  const exactRuleKey=t=>`EXACT|${norm(t.description)}|${norm(t.transactionType||t.paymentMethod)}|${Number(t.amount||0)}`
  const generalRuleKey=t=>`GENERAL|${norm(t.description)}|${norm(t.transactionType||t.paymentMethod)}`
  const findRule=t=>categoryRules[exactRuleKey(t)]||categoryRules[generalRuleKey(t)]||null

  const applyCategoryRules=rows=>rows.map(t=>{
    const rule=findRule(t)
    return rule?{...t,type:rule.type,category:rule.category,needsReview:false,ruleApplied:true,ruleScope:rule.scope}:t
  })
  const updatePreviewClassification=(id,field,value)=>{
    setImportPreview(prev=>prev.map(t=>t.id===id?{
      ...t,
      [field]:value,
      ...(field==='type'&&value==='Fixed'&&t.category==='기타'?{category:'주거'}:{}),
      needsReview:false,
      ruleApplied:false
    }:t))
  }
  const rememberPreviewRule=(t,scope='exact')=>{
    const key=scope==='general'?generalRuleKey(t):exactRuleKey(t)
    const next={...categoryRules,[key]:{type:t.type,category:t.category,scope}}
    setCategoryRules(next)
    localStorage.setItem('wonseok-finance-category-rules',JSON.stringify(next))
    setImportPreview(prev=>prev.map(x=>{
      const matches=scope==='general'?generalRuleKey(x)===key:exactRuleKey(x)===key
      return matches?{...x,type:t.type,category:t.category,needsReview:false,ruleApplied:true,ruleScope:scope}:x
    }))
  }

  const saveInstallments=next=>{
    setInstallments(next)
    localStorage.setItem('wonseok-finance-installments',JSON.stringify(next))
  }
  const monthDiff=(start,end)=>{
    const [sy,sm]=start.split('-').map(Number),[ey,em]=end.split('-').map(Number)
    return (ey-sy)*12+(em-sm)
  }
  const installmentStatus=item=>{
    const elapsed=monthDiff(item.startMonth,selectedMonth)
    let current=Math.max(0,Math.min(Number(item.months),elapsed+1))
    if(item.lastStatementMonth===selectedMonth&&item.lastKnownRound){current=Math.max(current,Number(item.lastKnownRound))}
    const monthly=item.billedPrincipal||Math.round(Number(item.totalAmount)/Number(item.months))
    return {current,monthly,remaining:Math.max(0,Number(item.months)-current),done:current>=Number(item.months)}
  }
  const addInstallment=e=>{
    e.preventDefault()
    if(!installmentForm.description.trim()||!Number(installmentForm.totalAmount)||!Number(installmentForm.months))return
    saveInstallments([...installments,{...installmentForm,totalAmount:Number(installmentForm.totalAmount),months:Number(installmentForm.months),id:Date.now()}])
    setInstallmentForm({...installmentForm,description:'',totalAmount:'',months:''})
  }
  const removeInstallment=id=>saveInstallments(installments.filter(x=>x.id!==id))

  const addInstallmentCandidate=c=>{
    const existing=installments.find(x=>x.sourceKey===c.key)
    const startMonth=existing?.startMonth||(()=>{
      if(c.currentRound===1)return c.statementDate.slice(0,7)
      const [y,m]=c.statementDate.slice(0,7).split('-').map(Number)
      const dt=new Date(y,m-1-(c.currentRound-1),1)
      return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}`
    })()
    const item={id:existing?.id||Date.now(),description:c.description,card:c.card,totalAmount:c.totalAmount,months:c.months,startMonth,sourceKey:c.key,lastKnownRound:c.currentRound,lastStatementMonth:c.statementDate.slice(0,7),billedPrincipal:c.billedPrincipal}
    saveInstallments(existing?installments.map(x=>x.id===existing.id?item:x):[...installments,item])
  }

  const normalizeDescription=v=>String(v||'').toLowerCase().replace(/\s+/g,'').replace(/[^0-9a-z가-힣]/g,'')
  const dateDistance=(a,b)=>Math.abs(Math.round((new Date(`${a}T00:00:00`)-new Date(`${b}T00:00:00`))/86400000))
  const findPossibleDuplicate=t=>{
    const amount=Number(t.amount)||0
    const desc=normalizeDescription(t.description)
    return transactions.find(saved=>{
      if(Number(saved.amount)!==amount||dateDistance(saved.date,t.date)>3)return false
      const savedDesc=normalizeDescription(saved.description)
      const sameDesc=desc&&savedDesc&&(desc===savedDesc||desc.includes(savedDesc)||savedDesc.includes(desc))
      const manualSource=String(saved.source||'').includes('수동')||saved.paymentMethod==='수동입력'
      return sameDesc||manualSource
    })||null
  }
  const duplicateInfo=t=>{
    const duplicate=findPossibleDuplicate(t)
    return duplicate?{duplicate,label:`기존 ${duplicate.date} · ${duplicate.description} · ${won(duplicate.amount)}`}:null
  }

  const ledgerKey=t=>`${t.date}|${t.description}|${Number(t.amount)}|${t.paymentMethod||'수동입력'}|${t.type}|${t.category}`

  const deleteLedgerItem=(item)=>{
    const label=`${item.date} · ${item.description} · ${won(item.amount)}`
    const ok=window.confirm(`대시보드 확정 내역에서 삭제할까요?\n\n${label}\n\n거래내역 작업공간은 변경되지 않습니다.`)
    if(!ok)return
    const next=ledger.filter(t=>t!==item)
    saveLedger(next)
  }

  const commitToLedger=()=>{
    if(!transactions.length){
      alert('대시보드에 반영할 거래내역이 없습니다.')
      return
    }
    const existing=new Set(ledger.map(ledgerKey))
    const fresh=transactions
      .filter(t=>!existing.has(ledgerKey(t)))
      .map((t,i)=>({...t,ledgerId:t.ledgerId||`ledger-${Date.now()}-${i}`,confirmedAt:new Date().toISOString()}))

    if(!fresh.length){
      alert('새로 반영할 거래가 없습니다. 이미 대시보드에 반영된 내역입니다.')
      return
    }

    saveLedger([...ledger,...fresh])
    alert(`${fresh.length}건을 대시보드 확정 원장에 반영했습니다.\n\n거래내역 작업공간을 비워도 대시보드 데이터는 유지됩니다.`)
  }

  const clearAllTransactions=()=>{
    if(!transactions.length)return
    const ok=window.confirm(`거래내역 ${transactions.length}건을 전부 삭제할까요?\n\n할부 현황과 분류 기억 규칙은 유지됩니다.\n이 작업은 되돌릴 수 없습니다.`)
    if(!ok)return
    save([])
    setSearch('')
    setTypeFilter('All')
    setMonthFilter('All')
    alert('거래내역 작업공간을 비웠습니다. 대시보드 확정 데이터는 유지됩니다.')
  }

  const exportBackup=()=>{
    const payload={
      app:'wonseok-finance-dashboard',
      version:APP_VERSION,
      exportedAt:new Date().toISOString(),
      transactions,
      ledger,
      installments,
      categoryRules
    }
    const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'})
    const url=URL.createObjectURL(blob)
    const a=document.createElement('a')
    a.href=url
    a.download=`wonseok-finance-backup-${new Date().toISOString().slice(0,10)}.json`
    document.body.appendChild(a)
    a.click()
    a.remove()
    URL.revokeObjectURL(url)
  }

  const importBackup=async e=>{
    const file=e.target.files?.[0]
    if(!file)return
    setBackupError('')
    try{
      const parsed=JSON.parse(await file.text())
      if(parsed.app!=='wonseok-finance-dashboard')throw new Error('이 앱의 백업 파일이 아닙니다.')
      if(!Array.isArray(parsed.transactions))throw new Error('거래 데이터가 없습니다.')

      save(parsed.transactions)

      const nextLedger=Array.isArray(parsed.ledger)?parsed.ledger:parsed.transactions
      saveLedger(nextLedger)

      const nextInstallments=Array.isArray(parsed.installments)?parsed.installments:[]
      setInstallments(nextInstallments)
      localStorage.setItem('wonseok-finance-installments',JSON.stringify(nextInstallments))

      const nextRules=parsed.categoryRules&&typeof parsed.categoryRules==='object'?parsed.categoryRules:{}
      setCategoryRules(nextRules)
      localStorage.setItem('wonseok-finance-category-rules',JSON.stringify(nextRules))

      alert(`백업 복원 완료\n거래 ${parsed.transactions.length}건 · 할부 ${nextInstallments.length}건`)
    }catch(err){
      setBackupError(`복원 실패: ${err.message}`)
    }
    e.target.value=''
  }

  const transactionKey=t=>`${t.date}|${t.description}|${t.amount}|${t.paymentMethod||'수동입력'}`
  const commitImport=()=>{
    if(!importPreview.length)return
    const existing=new Set(transactions.map(transactionKey))
    const exactFiltered=importPreview.filter(t=>!existing.has(transactionKey(t)))
    const skippedPossible=[]
    const fresh=[]
    exactFiltered.forEach(t=>{
      const dup=duplicateInfo(t)
      if(dup&&!duplicateOverrides[t.id])skippedPossible.push(t)
      else fresh.push({...t,id:Date.now()+fresh.length})
    })
    save([...transactions,...fresh]);setImportPreview([]);setImportName('');setImportStats(null);setDuplicateOverrides({})
    const exactSkipped=importPreview.length-exactFiltered.length
    alert([`${fresh.length}건을 거래내역에 추가했습니다.`,exactSkipped?`완전 중복 ${exactSkipped}건 제외`:'',skippedPossible.length?`중복 의심 ${skippedPossible.length}건 제외`:''].filter(Boolean).join('\n'))
  }

  return <main className="app">
    <header>
      <div><p>PERSONAL CFO</p><h1>원석 자산 대시보드</h1></div>
      <div className="headerActions">
        <nav>
          <button className={tab==='dashboard'?'active':''} onClick={()=>setTab('dashboard')}>대시보드</button>
          <button className={tab==='transactions'?'active':''} onClick={()=>setTab('transactions')}>거래내역</button>
          <button className={tab==='installments'?'active':''} onClick={()=>setTab('installments')}>할부 현황</button>
        </nav>
      </div>
    </header>

    {tab==='dashboard' ? <>
      <div className="monthBar">
        <div><span>조회 월</span><strong>{selectedMonth}</strong><em className="confirmedBadge">확정 데이터</em></div>
        <select value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)}>
          {availableMonths.length===0 && <option value="2026-07">2026-07</option>}
          {availableMonths.map(m=><option key={m} value={m}>{m.replace('-','년 ')}월</option>)}
        </select>
      </div>

      <section className="stats">
        <Stat label="총수입" value={won(totals.income)} Icon={WalletCards}/>
        <Stat label="순자산" value={won(totals.net)} Icon={Landmark}/>
        <Stat label="저축" value={won(totals.saving)} sub={`저축률 ${totals.income?(totals.saving/totals.income*100).toFixed(1):0}%`} Icon={PiggyBank}/>
        <Stat label="고정비" value={won(totals.fixed)} sub={`수입 대비 ${totals.income?(totals.fixed/totals.income*100).toFixed(1):0}%`} Icon={ReceiptText}/>
        <Stat label="변동비 가용예산" value={won(totals.budget)} Icon={Gauge}/>
        <Stat label="남은 변동비 예산" value={won(totals.remaining)} Icon={TrendingUp}/>
      </section>

      <section className="grid">
        <div className="panel"><h2>고정비 구성</h2><div className="rows">
          {fixedBy.length ? fixedBy.map(([c,a])=><div className="row" key={c}><span>{c}</span><strong>{won(a)}</strong><small>{totals.fixed?(a/totals.fixed*100).toFixed(1):0}%</small></div>)
          : <div className="empty">이 달의 고정비가 없습니다.</div>}
        </div></div>

        <div className="panel"><h2>월 현금흐름</h2><div className="rows">
          <div className="row2"><span>수입</span><strong>{won(totals.income)}</strong></div>
          <div className="row2"><span>저축</span><strong>- {won(totals.saving)}</strong></div>
          <div className="row2"><span>고정비</span><strong>- {won(totals.fixed)}</strong></div>
          <div className="row2"><span>변동비 사용액</span><strong>- {won(totals.variable)}</strong></div>
          <div className="row2 highlight"><span>남은 변동비 예산</span><strong>{won(totals.remaining)}</strong></div>
        </div></div>

        <div className="panel"><h2>변동비 카테고리</h2><div className="cards">
          {variableBy.map(([c,a])=><div className="mini" key={c}><span>{c}</span><strong>{won(a)}</strong></div>)}
        </div></div>

        <div className="panel"><h2>자산 / 부채</h2><div className="rows">
          <div className="row2"><span>총자산</span><strong>{won(ASSETS)}</strong></div>
          <div className="row2"><span>총부채</span><strong>{won(DEBTS)}</strong></div>
          <div className="row2 highlight"><span>순자산</span><strong>{won(totals.net)}</strong></div>
        </div></div>
      </section>

      <section className="panel confirmedLedgerPanel">
        <h2>대시보드 확정 내역 <small>{selectedMonth} · {monthTransactions.length}건</small></h2>
        <div className="confirmedLedgerList">
          {monthTransactions.length ? monthTransactions.map((t,i)=><div className="confirmedLedgerRow" key={t.ledgerId||`${t.id}-${t.date}-${i}`}>
            <div><strong>{t.description}</strong><small>{t.date} · {TYPE_LABELS[t.type]||t.type} · {t.category}</small></div>
            <b>{won(t.amount)}</b>
            <button type="button" className="ledgerDeleteButton" onClick={()=>deleteLedgerItem(t)} title="확정 내역에서 삭제"><Trash2 size={16}/><span>삭제</span></button>
          </div>) : <div className="empty">이 달의 확정 내역이 없습니다.</div>}
        </div>
      </section>
    </> : tab==='transactions' ? <>
      <section className="importPanel panel">
        <h2>카드 / 계좌 파일 업로드 <small>v{APP_VERSION}</small></h2>
        <div className="uploadBody">
          <label className="uploadButton"><Upload size={18}/> XLS / XLSX / CSV 선택<input type="file" accept=".xls,.xlsx,.csv" onChange={handleFile}/></label>
          <span className="uploadHint">미리보기에서 분류를 수정하고 ‘기억’을 누르면 같은 사용처는 다음 달부터 자동 분류됩니다.</span>
          {importName&&<div className="fileBadge"><FileSpreadsheet size={17}/><span>{importName}</span><b>{importPreview.length}건 감지</b><button onClick={()=>{setImportPreview([]);setImportName('');setImportError('');setImportStats(null);setInstallmentCandidates([])}}><X size={15}/></button></div>}
                    {importStats&&<div className="importStats">
            <span><b>{importStats.parser}</b></span>
            <span>가져올 거래 {importStats.imported}건</span>
            {(importStats.skippedInstallmentPayments+importStats.skippedSummaryRows)>0&&
              <span>집계 제외 {importStats.skippedInstallmentPayments+importStats.skippedSummaryRows}건</span>}
            {importStats.incomeSkipped>0&&<span>입금 제외 {importStats.incomeSkipped}건</span>}
          </div>}
          {installmentCandidates.length>0&&<div className="installmentDetected">
            <div className="installmentDetectedHead"><strong>할부 감지 {installmentCandidates.length}건</strong><span>소비에는 중복 반영하지 않고 할부 현황만 갱신합니다.</span></div>
            {installmentCandidates.map((c,i)=><div className="installmentDetectedRow" key={`${c.key}-${i}`}><div><strong>{c.description}</strong><small>{c.card} · {c.currentRound}/{c.months}회 · 이번 명세서 {won(c.billedPrincipal)}</small></div><b>{won(c.totalAmount)}</b><button onClick={()=>addInstallmentCandidate(c)}>할부 현황 반영</button></div>)}
          </div>}
          {importError&&<div className="importError">{importError}</div>}
          {importPreview.length>0&&<>
            <div className="previewHead"><strong>업로드 미리보기</strong><span>아직 저장되지 않았습니다.</span><button className="primary" onClick={commitImport}>이 {importPreview.length}건 추가</button></div>
            <div className="previewTable">{importPreview.map(t=>{const dup=duplicateInfo(t);const include=!!duplicateOverrides[t.id];return <div className={`previewRow previewRowEdit ${dup&&!include?'possibleDuplicate':''}`} key={t.id}>
              <span>{t.date}</span>
              <div className="previewMerchant">
                <strong>{t.description}</strong>
                <small>{t.paymentMethod}{t.ruleApplied?` · ${t.ruleScope==='exact'?'금액 규칙':'사용처 규칙'} 적용`:''}</small>
                {dup&&<small className="duplicateHint">중복 의심 · {dup.label}</small>}
              </div>
              <select value={t.type} onChange={e=>{
                  const type=e.target.value
                  updatePreviewClassification(t.id,'type',type)
                  updatePreviewClassification(t.id,'category',CATEGORY_MAP[type]?.[0]||'기타')
                }}>
                <option value="Variable">변동비</option>
                <option value="Fixed">고정비</option>
              </select>
              <select value={t.category} onChange={e=>updatePreviewClassification(t.id,'category',e.target.value)}>
                {(CATEGORY_MAP[t.type]||['기타']).map(c=><option key={c} value={c}>{c}</option>)}
              </select>
              <div className="rememberGroup">
                <button className={t.ruleApplied&&t.ruleScope==='exact'?'rememberBtn remembered':'rememberBtn'} onClick={()=>rememberPreviewRule(t,'exact')} title="같은 사용처·거래유형·금액에만 적용">{t.ruleApplied&&t.ruleScope==='exact'?'금액기억됨':'이 금액 기억'}</button>
                <button className={t.ruleApplied&&t.ruleScope==='general'?'rememberBtn remembered':'rememberBtn general'} onClick={()=>rememberPreviewRule(t,'general')} title="같은 사용처·거래유형 전체에 적용">{t.ruleApplied&&t.ruleScope==='general'?'사용처기억됨':'사용처 기억'}</button>
              </div>
              <div className="duplicateControl">{dup?<button type="button" className={include?'includeDuplicate included':'includeDuplicate'} onClick={()=>setDuplicateOverrides(v=>({...v,[t.id]:!v[t.id]}))}>{include?'추가함':'중복 제외'}</button>:(t.needsReview?<em>확인 필요</em>:<span className="reviewOk">확인</span>)}</div>
              <b>{won(t.amount)}</b>
            </div>})}</div>
            
          </>}
        </div>
      </section>
      <section className="ledgerCommitBar">
        <div>
          <strong>거래내역 작업공간</strong>
          <small>분류가 끝난 내역만 대시보드의 월별 확정 데이터로 반영하세요.</small>
        </div>
        <div className="ledgerCommitStats">
          <span>작업공간 <b>{transactions.length}건</b></span>
          <span>대시보드 확정 <b>{ledger.length}건</b></span>
        </div>
        <button type="button" className="commitLedgerButton" onClick={commitToLedger}>대시보드에 반영</button>
      </section>

      <section className="backupBar">
        <div>
          <DatabaseBackup size={18}/>
          <span><strong>데이터 백업</strong><small>작업공간·대시보드 확정 데이터·할부·분류 규칙을 한 파일로 저장합니다.</small></span>
        </div>
        <div className="backupActions">
          <button type="button" onClick={exportBackup}><Download size={16}/>백업 다운로드</button>
          <label className="backupUpload"><Upload size={16}/>백업 복원<input type="file" accept=".json,application/json" onChange={importBackup}/></label>
          <button type="button" className="dangerButton" onClick={clearAllTransactions}><Trash2 size={16}/>거래내역 작업공간 비우기</button>
        </div>
        {backupError&&<div className="backupError">{backupError}</div>}
      </section>

      <section className="transactionLayout">
        <div className="panel">
          <h2>{editingId?'거래 수정':'거래 추가'}</h2>
          <form onSubmit={submit}>
            <label>날짜<input type="date" value={form.date} onChange={e=>setForm({...form,date:e.target.value})}/></label>
            <label>내용<input value={form.description} onChange={e=>setForm({...form,description:e.target.value})} placeholder="예: 점심"/></label>
            <label>금액<input type="number" value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})} placeholder="12000"/></label>
            <label>구분<select value={form.type} onChange={e=>{const type=e.target.value;setForm({...form,type,category:MANUAL_CATEGORIES[type]?.[0]||'기타'})}}>
              {Object.entries(TYPE_LABELS).map(([v,l])=><option key={v} value={v}>{l}</option>)}
            </select></label>
            <label>카테고리<select list="cats" value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>
                {(MANUAL_CATEGORIES[form.type]||['기타']).map(c=><option key={c} value={c}>{c}</option>)}
              </select></label>
            <datalist id="cats">{VAR_CATS.map(c=><option key={c} value={c}/>)}</datalist>
            <button className="primary" type="submit"><Plus size={16}/>{editingId?'수정 저장':'추가'}</button>
            {editingId&&<button type="button" onClick={()=>{setEditingId(null);setForm({...form,description:'',amount:''})}}>취소</button>}
          </form>
        </div>

        <div className="panel">
          <h2>거래내역 <small>{filteredTransactions.length}건</small></h2>
          <div className="filters">
            <div className="searchBox"><Search size={17}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="내용 또는 카테고리 검색"/></div>
            <select value={monthFilter} onChange={e=>setMonthFilter(e.target.value)}>
              <option value="All">전체 월</option>
              {importMonths.map(m=><option key={m} value={m}>{m}</option>)}
            </select>
            <select value={typeFilter} onChange={e=>setTypeFilter(e.target.value)}>
              <option value="All">전체 구분</option>
              {Object.entries(TYPE_LABELS).map(([v,l])=><option key={v} value={v}>{l}</option>)}
            </select>
            <button className="iconText" onClick={resetFilters}><RotateCcw size={16}/>필터 초기화</button>
          </div>
          <div className="transactionList">
            {filteredTransactions.length ? filteredTransactions.map(t=><div className="transaction" key={t.id}>
              <div><strong>{t.description}</strong><small>{t.date} · {TYPE_LABELS[t.type]} · {t.category} · {t.paymentMethod||'수동입력'}</small></div>
              <b>{won(t.amount)}</b>
              <button onClick={()=>edit(t)}><Pencil size={16}/></button>
              <button onClick={()=>remove(t.id)}><Trash2 size={16}/></button>
            </div>) : <div className="empty">조건에 맞는 거래가 없습니다.</div>}
          </div>
        </div>
      </section>
    </> : <>
      <div className="monthBar">
        <div><span>기준 월</span><strong>{selectedMonth}</strong></div>
        <select value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)}>
          {availableMonths.map(m=><option key={m} value={m}>{m}</option>)}
        </select>
      </div>
      <section className="installmentSummary">
        <div className="stat"><span>진행 중 할부</span><strong>{installments.filter(x=>!installmentStatus(x).done).length}건</strong></div>
        <div className="stat"><span>이번 달 예상 할부액</span><strong>{won(installments.reduce((a,x)=>a+(installmentStatus(x).done?0:installmentStatus(x).monthly),0))}</strong></div>
        <div className="stat"><span>남은 할부 원금</span><strong>{won(installments.reduce((a,x)=>{const st=installmentStatus(x);return a+st.monthly*st.remaining},0))}</strong></div>
      </section>
      <section className="transactionLayout installmentLayout">
        <div className="panel">
          <h2>할부 추가</h2>
          <form onSubmit={addInstallment}>
            <label>구매 항목<input value={installmentForm.description} onChange={e=>setInstallmentForm({...installmentForm,description:e.target.value})} placeholder="예: 카메라"/></label>
            <label>카드<select value={installmentForm.card} onChange={e=>setInstallmentForm({...installmentForm,card:e.target.value})}><option>현대카드</option><option>국민카드</option><option>기타</option></select></label>
            <label>총 구매금액<input type="number" value={installmentForm.totalAmount} onChange={e=>setInstallmentForm({...installmentForm,totalAmount:e.target.value})} placeholder="600000"/></label>
            <label>할부 개월<input type="number" min="2" value={installmentForm.months} onChange={e=>setInstallmentForm({...installmentForm,months:e.target.value})} placeholder="6"/></label>
            <label>첫 결제월<input type="month" value={installmentForm.startMonth} onChange={e=>setInstallmentForm({...installmentForm,startMonth:e.target.value})}/></label>
            <button className="primary" type="submit"><Plus size={16}/>할부 추가</button>
          </form>
        </div>
        <div className="panel">
          <h2>할부 현황 <small>{installments.length}건</small></h2>
          <div className="installmentList">
            {installments.length?installments.map(x=>{const st=installmentStatus(x);return <div className="installmentCard" key={x.id}>
              <div className="installmentTop"><div><strong>{x.description}</strong><small>{x.card} · {x.startMonth} 시작{x.lastStatementMonth?` · 최근 명세서 ${x.lastStatementMonth} (${x.lastKnownRound}/${x.months})`:''}</small></div><button onClick={()=>removeInstallment(x.id)}><Trash2 size={16}/></button></div>
              <div className="installmentNumbers"><span>구매금액 <b>{won(x.totalAmount)}</b></span><span>월 예상 <b>{won(st.monthly)}</b></span><span>진행 <b>{st.done?'완료':`${st.current}/${x.months}회`}</b></span><span>남은 회차 <b>{st.remaining}회</b></span></div>
              <div className="progress"><i style={{width:`${Math.min(100,(st.current/x.months)*100)}%`}}/></div>
            </div>}):<div className="empty">등록된 할부가 없습니다. 카드 명세서에서 기억하기 어려운 할부만 추가하면 됩니다.</div>}
          </div>
        </div>
      </section>
    </>}
  </main>
}
