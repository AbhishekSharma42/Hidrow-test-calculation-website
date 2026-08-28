import React,{useMemo,useState} from "react";
import {createRoot} from "react-dom/client";
import {Calculator,FileText,Activity,Droplets,CheckCircle2,AlertTriangle} from "lucide-react";
import "./styles.css";

const fmt=(n,d=4)=>Number.isFinite(n)?Number(n).toFixed(d):"—";
const round=(n,d=4)=>Number(Number(n).toFixed(d));

/* The workbook's A-value table: A = base - coefficient * P / 300.
   Values are taken from the supplied calculations hydro.xlsx. */
const aRows=[
 [2,49.45,2.81],[4,48.79,2.79],[6,48.23,2.83],[8,47.71,2.78],
 [10,47.27,2.83],[12,46.88,2.89],[14,46.48,2.96],[16,46.10,2.93],
 [18,45.82,3.05],[20,45.49,2.96],[22,45.25,3.03],[24,45.06,3.08],
 [26,44.87,3.12],[28,44.70,3.15],[30,44.58,3.24]
];
function interp(x,x1,y1,x2,y2){if(x2===x1)return y1;return y1+(y2-y1)*(x-x1)/(x2-x1)}
function aValue(temp,pressure){
  const t=Math.max(2,Math.min(30,temp));
  let i=aRows.findIndex(r=>t<=r[0]);
  if(i<=0){const r=aRows[0];return r[1]-r[2]*pressure/300}
  const lo=aRows[i-1],hi=aRows[i];
  const a1=lo[1]-lo[2]*pressure/300,a2=hi[1]-hi[2]*pressure/300;
  return interp(t,lo[0],a1,hi[0],a2);
}

/* Thermal expansion B lookup from the workbook's Strength & Leak Test
   table. The table is pressure (10..250 bar) x temperature (1..34 °C).
   For the normal test range we reproduce the sheet by bilinear interpolation. */
const pressureAxis=[10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250];
const bBase=[
 [-75.03,-57.25,-40.10,-23.54,-7.56,7.92,22.89,37.89,51.44,65.08,78.32,91.17,103.68,115.84,127.69,139.22,150.49,161.46,172.18,182.66,192.91,202.93,212.74,222.35,231.78,241.02,250.08,258.99,267.75,276.35,284.96,293.38,301.64,309.74],
 [-70.40,-53.96,-36.94,-20.51,-4.65,10.70,25.55,39.94,53.88,67.40,80.53,93.29,105.69,117.76,129.50,140.95,152.11,163.00,173.64,184.03,194.20,204.14,213.88,223.42,232.77,241.94,250.93,259.78,268.47,277.01,285.33,293.42,301.31,308.98],
 [-65.76,-50.66,-33.77,-17.47,-1.73,13.50,28.23,42.50,56.32,69.73,82.75,95.41,107.70,119.67,131.32,142.67,153.75,164.56,175.10,185.41,195.50,205.36,215.03,224.48,233.76,242.87,251.79,260.57,269.19,277.66,285.84,293.72,301.35,308.78],
 [-61.28,-44.05,-27.43,-11.38,4.10,19.08,33.58,47.67,61.21,74.41,87.22,99.66,111.75,123.52,134.98,146.15,157.04,167.66,178.05,188.20,198.12,207.82,217.33,226.64,235.78,244.73,253.53,262.17,270.77,279.01,287.01,294.75,302.29,309.70],
 [-57.84,-40.74,-24.26,-8.34,7.02,21.88,36.26,50.18,63.67,76.74,89.45,101.79,113.79,125.46,136.82,147.90,158.70,169.24,179.54,189.59,199.44,209.06,218.49,227.73,236.78,245.68,254.40,262.98,271.41,279.69,287.82,295.80,303.62,311.30],
 [-54.40,-37.44,-21.08,-5.29,9.95,24.68,38.94,52.75,66.12,79.09,91.69,103.93,115.83,127.39,138.67,149.65,160.36,170.81,181.02,191.00,200.75,210.31,219.66,228.82,237.81,246.63,255.28,263.69,272.16,280.38,288.49,296.48,304.32,311.98],
 [-50.96,-34.13,-17.90,-2.25,12.87,27.49,41.63,55.32,68.58,81.45,93.93,106.07,117.87,129.34,140.51,151.40,162.03,172.39,182.51,192.41,202.09,211.56,220.84,229.92,238.84,247.59,256.18,264.62,272.92,281.08,289.17,297.12,304.94,312.61],
 [-47.53,-30.83,-14.73,0.80,15.79,30.29,44.31,57.89,71.05,83.80,96.18,108.21,119.90,131.20,142.37,153.16,163.70,173.98,184.00,193.82,203.42,212.81,222.02,231.03,239.87,248.55,257.07,265.44,273.69,281.78,289.84,297.78,305.58,313.22],
 [-44.10,-27.53,-11.56,3.85,18.72,33.10,47.00,60.46,73.51,86.15,98.43,110.36,121.96,133.74,144.22,154.93,165.37,175.56,185.51,195.24,204.76,214.08,223.20,232.14,240.91,249.52,257.97,266.28,274.46,282.49,290.49,298.42,306.19,313.82],
 [-40.67,-24.23,-8.40,6.89,21.65,35.90,49.69,63.04,75.97,88.51,100.68,112.51,124.01,135.19,146.08,156.69,167.05,177.15,187.02,196.66,206.10,215.34,224.39,233.26,241.96,250.49,258.88,267.12,275.23,283.20,291.15,299.03,306.77,314.36],
 [-37.24,-20.94,-5.23,9.94,24.56,38.70,52.37,65.62,78.44,90.87,102.94,114.66,126.06,137.15,147.94,158.47,168.73,178.75,188.53,198.09,207.45,216.61,225.58,234.38,243.01,251.47,259.79,267.97,276.01,283.92,291.82,299.64,307.34,314.91],
 [-33.83,-17.65,-2.06,12.98,27.48,41.51,55.06,68.19,80.91,93.23,105.19,116.82,128.12,139.11,149.81,160.24,170.42,180.35,190.05,199.52,208.80,217.89,226.79,235.51,244.06,252.46,260.71,268.82,276.80,284.64,292.44,300.18,307.85,315.39]
];
function bValue(temp,pressure){
  const p=Math.max(10,Math.min(100,pressure));
  const t=Math.max(1,Math.min(34,temp));
  let pi=pressureAxis.findIndex(x=>p<=x); if(pi<0)pi=pressureAxis.length-1;
  if(pi===0)pi=1;
  const p1=pressureAxis[pi-1],p2=pressureAxis[pi];
  const row1=bBase[pi-1]||bBase[bBase.length-1],row2=bBase[pi]||bBase[bBase.length-1];
  const tempIndex=t-1; const ti=Math.floor(tempIndex),tj=Math.min(33,ti+1),f=tempIndex-ti;
  const v11=row1[ti],v12=row1[tj],v21=row2[ti],v22=row2[tj];
  const v1=v11+(v12-v11)*f, v2=v21+(v22-v21)*f;
  return v1+(v2-v1)*(p-p1)/(p2-p1);
}

function Intro({data,setData}){
 const rows=data.rows;
 const update=(i,k,v)=>setData({...data,rows:rows.map((r,j)=>j===i?{...r,[k]:v}:r)});
 const calc=useMemo(()=>{
   let totalL=0,totalV=0,wt=0,ri=0;
   rows.forEach(r=>{const od=+r.od||0,t=+r.t||0,l=+r.length||0,inner=Math.max(0,od-2*t);const vol=Math.PI/4*inner*inner*l/1e6;totalL+=l;totalV+=vol;wt+=t*l;ri+=inner*l});
   const avgT=totalL?wt/totalL:0,avgInner=totalL?ri/totalL:0;
   return {avgT,avgRadius:avgInner/2,avgVolPerM:totalL?totalV/totalL:0,totalL,totalV,ratio:avgT?avgInner/2/avgT:0};
 },[rows]);
 return <Page title="HT Section Introduction" subtitle="Pipe section inputs and automatic volume calculations">
   <div className="grid">
    <div className="card span-12"><h3 className="section-title">Section Details</h3><div className="fields">
      <div className="field"><label>From Joint</label><input value={data.fromJoint} onChange={e=>setData({...data,fromJoint:e.target.value})}/></div>
      <div className="field"><label>To Joint</label><input value={data.toJoint} onChange={e=>setData({...data,toJoint:e.target.value})}/></div>
      <div className="field"><label>From Chainage</label><div className="unit-input"><input type="number" value={data.fromChainage} onChange={e=>setData({...data,fromChainage:e.target.value})}/><span className="unit">m</span></div></div>
      <div className="field"><label>To Chainage</label><div className="unit-input"><input type="number" value={data.toChainage} onChange={e=>setData({...data,toChainage:e.target.value})}/><span className="unit">m</span></div></div>
    </div></div>
    <div className="card span-12"><h3 className="section-title">Pipe Segments</h3><div className="table-wrap"><table className="table"><thead><tr><th>Sr.</th><th>Outer Dia (mm)</th><th>Wall Thick (mm)</th><th>Inner Dia (mm)</th><th>Length (m)</th><th>Volume (m³)</th></tr></thead><tbody>
      {rows.map((r,i)=>{const inner=Math.max(0,(+r.od||0)-2*(+r.t||0));const vol=Math.PI/4*inner*inner*(+r.length||0)/1e6;return <tr key={i}><td>{i+1}</td><td><input type="number" value={r.od} onChange={e=>update(i,"od",e.target.value)}/></td><td><input type="number" value={r.t} onChange={e=>update(i,"t",e.target.value)}/></td><td>{fmt(inner,3)}</td><td><input type="number" value={r.length} onChange={e=>update(i,"length",e.target.value)}/></td><td>{fmt(vol,6)}</td></tr>})}
    </tbody></table></div></div>
    <div className="card span-12"><h3 className="section-title">Calculated Output</h3><div className="results">
      <Metric label="Avg. Wall Thickness" value={fmt(calc.avgT,4)} unit="mm"/><Metric label="Avg. Inner Radius" value={fmt(calc.avgRadius,4)} unit="mm"/>
      <Metric label="Avg. Volume / m" value={fmt(calc.avgVolPerM,8)} unit="m³"/><Metric label="Total Length" value={fmt(calc.totalL,2)} unit="m"/>
      <Metric label="Total Section Volume" value={fmt(calc.totalV,8)} unit="m³"/><Metric label="Ri / t Ratio" value={fmt(calc.ratio,6)} unit=""/>
    </div></div>
   </div>
 </Page>
}

function Air({intro,data,setData}){
 const [saved,setSaved]=useState(false);
 const calc=useMemo(()=>{
   const A=aValue(+data.temp,+data.pressure);
   const theoretical=((0.884*intro.avgRadius/intro.avgT)+A)*1e-6*intro.totalV*(+data.deltaP)*+data.K;
   const liters=theoretical*1000; const ratio=(+data.actualLiters||0)/liters;
   return {A,theoretical,liters,ratio,accepted:ratio<=1.02};
 },[intro,data]);
 return <Page title="Air Volume Calculation" subtitle="Calculation matching the supplied workbook logic">
  <div className="grid">
   <div className="card span-7"><h3 className="section-title">Test Inputs</h3><div className="fields">
    <Field label="Pipe Temp." value={data.temp} unit="°C" onChange={v=>setData({...data,temp:v})}/>
    <Field label="Pressure (Dead Weight)" value={data.pressure} unit="bar" onChange={v=>setData({...data,pressure:v})}/>
    <Field label="Dimensional Coefficient K" value={data.K} unit="" onChange={v=>setData({...data,K:v})}/>
    <Field label="Change of Pressure ΔP" value={data.deltaP} unit="bar" onChange={v=>setData({...data,deltaP:v})}/>
    <Field label="Actual Volume of Water VA" value={data.actualLiters} unit="L" onChange={v=>setData({...data,actualLiters:v})}/>
   </div><div className="btn-row"><button className="btn btn-primary" onClick={()=>setSaved(true)}>Calculate</button><button className="btn btn-secondary" onClick={()=>setData({...data,temp:27.5,pressure:98.02,K:1.02,deltaP:.5,actualLiters:431.2})}>Reset Workbook Values</button></div>{saved&&<p className="note">Calculation updated live. The original workbook's sample values are loaded by default.</p>}</div>
   <div className="card span-5"><h3 className="section-title">Calculated Output</h3><div className="results">
    <Metric label="Isothermal Compressibility A" value={fmt(calc.A,6)} unit=""/>
    <Metric label="Inner Pipe Radius Ri" value={fmt(intro.avgRadius,4)} unit="mm"/>
    <Metric label="Pipe Thickness t" value={fmt(intro.avgT,4)} unit="mm"/>
    <Metric label="Total Volume Vt" value={fmt(intro.totalV,6)} unit="m³"/>
    <Metric label="Theoretical Water VP" value={fmt(calc.theoretical,9)} unit="m³"/>
    <Metric label="Theoretical Water VP" value={fmt(calc.liters,6)} unit="L"/>
    <Metric label="Air Volume Ratio" value={fmt(calc.ratio,6)} unit=""/>
   </div><div className={"status "+(calc.accepted?"pass":"fail")}>{calc.accepted?<><CheckCircle2 size={18} style={{verticalAlign:"middle",marginRight:6}}/>Air Volume Ratio Accepted</>:<><AlertTriangle size={18} style={{verticalAlign:"middle",marginRight:6}}/>Air Volume Ratio is not Accepted</>}</div></div>
   <div className="card span-12"><h3 className="section-title">Workbook Formula</h3><div className="formula">VP = (0.884 × Ri / t + A) × 10⁻⁶ × Vt × ΔP × K</div><p className="note">A is interpolated from the temperature/pressure table in the supplied workbook. Values above 30 °C can be added to the same table if your project uses that range.</p></div>
  </div>
 </Page>
}

function Strength({intro,air,data,setData}){
 const calc=useMemo(()=>{
  const dT=(+data.initialTemp)-(+data.finalTemp);
  const dP=(+data.finalPressure)-(+data.initialPressure);
  const A=aValue(+data.finalTemp,+data.finalPressure);
  const B=bValue(+data.finalTemp,+data.finalPressure);
  const accountable=B*dT/((0.884*intro.ratio)+A);
  const loss=Math.abs(dP)-accountable;
  const pass=loss<=.3;
  return {dT,dP,A,B,accountable,loss,pass};
 },[intro,data]);
 return <Page title="Strength & Leak Test" subtitle="Temperature correction, pressure loss and final pass/fail result">
  <div className="grid">
   <div className="card span-6"><h3 className="section-title">Initial / Final Test Data</h3><div className="fields">
    <Field label="Initial Pipe Wall Temp. Ti" value={data.initialTemp} unit="°C" onChange={v=>setData({...data,initialTemp:v})}/>
    <Field label="Final Pipe Wall Temp. Tf" value={data.finalTemp} unit="°C" onChange={v=>setData({...data,finalTemp:v})}/>
    <Field label="Initial Test Pressure Pi" value={data.initialPressure} unit="bar" onChange={v=>setData({...data,initialPressure:v})}/>
    <Field label="Final Test Pressure Pf" value={data.finalPressure} unit="bar" onChange={v=>setData({...data,finalPressure:v})}/>
   </div></div>
   <div className="card span-6"><h3 className="section-title">Calculated Result</h3><div className="results">
    <Metric label="Change of Temperature ΔT" value={fmt(calc.dT,3)} unit="°C"/>
    <Metric label="Change of Pressure ΔP" value={fmt(calc.dP,3)} unit="bar"/>
    <Metric label="Isothermal Compressibility A" value={fmt(calc.A,6)} unit=""/>
    <Metric label="B — thermal expansion difference" value={fmt(calc.B,6)} unit=""/>
    <Metric label="Accountable Pressure Change ΔPt" value={fmt(calc.accountable,6)} unit="bar"/>
    <Metric label="Unaccountable Pressure Loss" value={fmt(calc.loss,6)} unit="bar"/>
   </div><div className={"status "+(calc.pass?"pass":"fail")}>{calc.pass?<><CheckCircle2 size={18} style={{verticalAlign:"middle",marginRight:6}}/>The Strength/Leak Test for the Entire Section is Successfully Pass</>:<><AlertTriangle size={18} style={{verticalAlign:"middle",marginRight:6}}/>The Entire Section is Fail in Strength/Leak Test</>}</div></div>
   <div className="card span-12"><h3 className="section-title">Acceptance</h3><p className="note">Workbook rule: allowable pressure loss is ≤ 0.3 bar after temperature correction. The B value is obtained by interpolation from the supplied pressure/temperature table.</p><div className="formula">ΔPt = B × ΔT / ((0.884 × Ri/t) + A){"\n"}Unaccountable Pressure Loss = ΔP − ΔPt</div></div>
  </div>
 </Page>
}

function Dashboard({intro,air,strength}){
 return <Page title="Dashboard" subtitle="Abhi Sharma Hydro-Test Calculator — workbook overview">
  <div className="grid">
   <div className="card span-4"><h3 className="section-title"><Droplets size={18} style={{verticalAlign:"middle"}}/> Section Volume</h3><div className="metric"><div className="label">Total Volume</div><div className="value">{fmt(intro.totalV,6)}</div><div className="unit2">m³</div></div></div>
   <div className="card span-4"><h3 className="section-title"><Activity size={18} style={{verticalAlign:"middle"}}/> Air Volume Ratio</h3><div className="metric"><div className="label">Current Ratio</div><div className="value">{fmt(air.ratio,4)}</div><div className="unit2">limit ≤ 1.02</div></div></div>
   <div className="card span-4"><h3 className="section-title"><FileText size={18} style={{verticalAlign:"middle"}}/> Strength / Leak</h3><div className={"status "+(strength.pass?"pass":"fail")} style={{marginTop:0}}>{strength.pass?"PASS":"FAIL"}</div></div>
   <div className="card span-12"><h3 className="section-title">Workbook Pages</h3><div className="grid">
    {["HT Sec. Intro","Air Volume Calculation","Strength & Leak Test"].map((x,i)=><div className="card span-4" key={x}><strong>{i+1}. {x}</strong><p className="note">{i===0?"Pipe section dimensions, average thickness/radius and total volume.":i===1?"Theoretical water volume, actual water volume and air volume ratio.":"Temperature correction, accountable pressure change and final test result."}</p></div>)}
   </div></div>
  </div>
 </Page>
}
function Metric({label,value,unit}){return <div className="metric"><div className="label">{label}</div><div className="value">{value}</div><div className="unit2">{unit}</div></div>}
function Field({label,value,unit,onChange}){return <div className="field"><label>{label}</label><div className="unit-input"><input type="number" value={value} onChange={e=>onChange(e.target.value)}/>{unit&&<span className="unit">{unit}</span>}</div></div>}
function Page({title,subtitle,children}){return <><div className="hero"><h2>{title}</h2><p>{subtitle}</p></div>{children}</>}
function App(){
 const [page,setPage]=useState("dashboard");
 const [intro,setIntro]=useState({fromJoint:"0.0 Km FT09",toJoint:"21.0 Km FT23",fromChainage:0,toChainage:20871.33,rows:[
  {od:114.3,t:6.4,length:1500},{od:"",t:"",length:""},{od:"",t:"",length:""},{od:"",t:"",length:""},{od:"",t:"",length:""},{od:"",t:"",length:0},{od:"",t:"",length:0}
 ]});
 const [air,setAir]=useState({temp:27.5,pressure:98.02,K:1.02,deltaP:.5,actualLiters:431.2});
 const [strength,setStrength]=useState({initialTemp:26.3,finalTemp:25.6,initialPressure:75.06,finalPressure:72.1});
 const introCalc=useMemo(()=>{let L=0,V=0,wt=0,ri=0;intro.rows.forEach(r=>{const od=+r.od||0,t=+r.t||0,l=+r.length||0,inn=Math.max(0,od-2*t);L+=l;V+=Math.PI/4*inn*inn*l/1e6;wt+=t*l;ri+=inn*l});const avgT=L?wt/L:0,avgInner=L?ri/L:0;return {avgT,avgRadius:avgInner/2,avgVolPerM:L?V/L:0,totalL:L,totalV:V,ratio:avgT?avgInner/2/avgT:0}},[intro]);
 const airCalc=useMemo(()=>{const A=aValue(+air.temp,+air.pressure),th=((.884*introCalc.avgRadius/introCalc.avgT)+A)*1e-6*introCalc.totalV*(+air.deltaP)*+air.K,lit=th*1000,ratio=(+air.actualLiters||0)/lit;return {A,theoretical:th,liters:lit,ratio,accepted:ratio<=1.02}},[introCalc,air]);
 const strengthCalc=useMemo(()=>{const dT=+strength.initialTemp-+strength.finalTemp,dP=+strength.finalPressure-+strength.initialPressure,A=aValue(+strength.finalTemp,+strength.finalPressure),B=bValue(+strength.finalTemp,+strength.finalPressure),accountable=B*dT/((.884*introCalc.ratio)+A),loss=Math.abs(dP)-accountable;return {dT,dP,A,B,accountable,loss,pass:loss<=.3}},[introCalc,strength]);
 const nav=[["dashboard","Dashboard"],["intro","HT Sec. Intro"],["air","Air Volume"],["strength","Strength & Leak"]];
 return <div className="app"><header className="topbar"><div className="brand"><div className="brand-mark">AS</div><div><h1>Abhi Sharma</h1><small>Hydro-Test Calculator</small></div></div><nav className="nav">{nav.map(([id,label])=><button className={page===id?"active":""} onClick={()=>setPage(id)} key={id}>{label}</button>)}</nav></header><main className="container">
 {page==="dashboard"&&<Dashboard intro={introCalc} air={airCalc} strength={strengthCalc}/>}
 {page==="intro"&&<Intro data={intro} setData={setIntro}/>}
 {page==="air"&&<Air intro={introCalc} data={air} setData={setAir}/>}
 {page==="strength"&&<Strength intro={introCalc} air={airCalc} data={strength} setData={setStrength}/>}
 </main><footer className="footer">© 2026 Abhi Sharma • Built from the structure and calculation logic of the supplied Excel workbook.</footer></div>
}
createRoot(document.getElementById("root")).render(<App/>);
