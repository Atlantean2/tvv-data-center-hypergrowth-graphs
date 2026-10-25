"""Builds TechVestVentures Data Center Hypergrowth Graphs from focus/data-center (8 PNGs; titles/source lines as drawn on the charts by scripts/build.py; bullets verbatim from the Risk Brief docx)."""
import os,sys,shutil,docx
from docx.oxml.ns import qn
sys.path.insert(0,'/workspace/risk-watcher/graphs-site');os.chdir(os.path.dirname(os.path.abspath(__file__)))
import common
from common import *
T="TechVestVentures Data Center Hypergrowth Graphs";F="/workspace/risk-watcher/focus/data-center"
D=docx.Document(F+"/TVV Data Center Financing Risk Brief 2026-10-25.docx");PS=D.paragraphs
def links(p): return [D.part.rels[h.get(qn('r:id'))].target_ref for h in p._p.iter(qn('w:hyperlink')) if h.get(qn('r:id'))]
def bul(i): p=PS[i];us=links(p);t=p.text.removesuffix("(source)").strip(); return f'<li>{E(t)} '+" ".join(f'<a href="{E(u)}">(source)</a>' for u in us)+'</li>'
# file, title (chart title), chart source line (verbatim), source urls, brief bullet paragraph indexes
C=[("01_iea_dc_electricity","Data-center electricity demand roughly doubles by 2030","IEA, Energy and AI (2025). US/China 2024 derived from IEA shares (45%/25%) + stated increments.",["https://www.iea.org/reports/energy-and-ai/executive-summary"],[10]),
("02_epoch_capacity_growth","Frontier AI data centers tracked by Epoch: power built & planned","Epoch AI, Frontier Data Centers Hub (timelines, satellite/permit-based; post-today = planned).",["https://epoch.ai/data/data-centers"],[10]),
("03a_map_us_campuses","Major AI data-center campuses — United States (bubble = MW; red = project-level debt/lease financing disclosed)","Epoch AI; aidatacenterindex.com; company filings/press (see xlsx). Locations approximate, hand-geocoded.",["https://epoch.ai/data/data-centers","https://aidatacenterindex.com/largest/"],[12,13,14]),
("03b_map_global_campuses","Major AI data-center campuses — global (bubble = MW; red = project-level debt/lease financing disclosed)","Epoch AI; aidatacenterindex.com; company filings/press (see xlsx). Locations approximate, hand-geocoded.",["https://epoch.ai/data/data-centers","https://aidatacenterindex.com/largest/"],[]),
("04_financing_mix","$131B of USD AI data-center debt, Jan 2025 – Oct 7 2026 (45 instruments)","ai-industry-reviews.com AI data center debt financing statistics 2026 (SEC 8-K based); excludes hyperscaler corporate bonds.",["https://ai-industry-reviews.com/research/ai-data-center-debt-financing-statistics-2026/"],[7,15,16,23]),
("05_debt_by_borrower","Who borrowed: hyperscaler bonds (blue) vs. neocloud/SPV/miner debt (red)","ai-industry-reviews.com (SEC filings), Jan 2025–Oct 2026.",["https://ai-industry-reviews.com/research/ai-data-center-debt-financing-statistics-2026/"],[7,8]),
("06_uncommenced_leases","Off-balance-sheet lease commitments: ~$1.2T and rising","Company 10-K/10-Q (ASC 842) via ahasignals.com, ccir.io; Oracle 10-Q Aug-26; WSJ Aug 2026 analysis.",["https://ahasignals.com/research/uncommenced-lease-commitments/","https://ccir.io/credit/commitments","https://www.sec.gov/Archives/edgar/data/1341439/000119312526389274/orcl-20260831.htm","https://www.wsj.com/tech/ai/why-big-techs-ai-spending-is-3-trillion-higher-than-it-seems-e1067bb2"],[3,4,5,22,27]),
("07_spreads_cds","Credit pricing of AI data-center risk (Oct 2026): bonds (blue), CDS (orange)","Bloomberg via aiinfracredit.com; securitasglobal.com; CNBC Oct 19 2026; Globe and Mail. Ratings: S&P.",["https://www.aiinfracredit.com/en/research/which-ai-debt-is-repricing","https://www.securitasglobal.com/risk-perspectives/oracle-cds-record-288-basis-points-ai-debt/","https://www.cnbc.com/2026/10/19/oracle-debt-faces-an-ai-stress-test-why-some-investors-arent-worried-.html","https://www.theglobeandmail.com/investing/article-are-ai-credit-cracks-a-warning-or-a-buy-signal/"],[9,19,20,21,25])]
C={c[0]:c for c in C}
TABS=[("Q1","Capacity Growth",["01_iea_dc_electricity","02_epoch_capacity_growth"]),("Q2","Maps",["03a_map_us_campuses","03b_map_global_campuses"]),
("Q3","Financing",["04_financing_mix","05_debt_by_borrower"]),("Q4","Off-Balance-Sheet Exposure",["06_uncommenced_leases"]),("Q5","Credit Stress",["07_spreads_cds"])]
common.CSS+="""
figure.g{display:grid;grid-template-columns:minmax(0,7fr) minmax(260px,3fr);gap:22px;align-items:start;margin:0 0 40px}
figure.g img{width:100%;height:auto;max-width:none}
figure.g figcaption{font-size:18px;line-height:1.5;margin:0}figure.g figcaption p{margin:0 0 12px}
figure.g figcaption .t{font-weight:700;display:block;margin-bottom:8px}figure.g ul{margin:0 0 12px;padding-left:1.1em;font-size:16px}figure.g li{margin:4px 0}
.view{border-left:4px solid #8B0000;background:#faf3f3;padding:10px 14px;margin:24px 0;font-size:17px;line-height:1.5}
@media(max-width:1100px){figure.g{grid-template-columns:1fr}}"""
shutil.rmtree("issues",ignore_errors=True);os.makedirs("issues/img");tabs=[]
def w(k,lab,body): open(f"issues/{k}.html","w").write(page(T,k,body));tabs.append((k,lab))
assess=next(p.text for p in PS if p.text.startswith("Assessment:"))
for q,lab,fs in TABS:
    figs=[]
    for f in fs:
        _,t,sl,us,bs=C[f];save_img(f"{F}/charts/{f}.png",f"issues/img/{f}.webp")
        b="<ul>"+"".join(bul(i) for i in bs)+"</ul>" if bs else ""
        figs.append(f'<figure class="g"><img src="img/{f}.webp" alt="{E(t)}" loading="lazy"><figcaption><span class="t">{E(t)}</span>{b}<p>Source: {E(sl)}<br>'+" ".join(f'<a href="{E(u)}">{E(u)}</a>' for u in us)+'</p></figcaption></figure>')
    extra=f'<div class="view">{E(assess)}</div>' if q=="Q5" else ""
    w(q,f"{q} {lab}",f'<h2 class="desk">{q}. {E(lab)}</h2>'+"\n".join(figs)+extra)
i=[k for k,p in enumerate(PS) if p.text=="Sources"][-1];src=[p.text for p in PS[i+1:] if p.text.startswith("https://")]
w("sources","Sources",f'<h2 class="desk">Sources ({len(src)} links)</h2><p>From the TVV Data Center Financing Risk Brief, October 25, 2026.</p><ul class="src">'+"".join(f'<li><a href="{E(u)}">{E(u)}</a></li>' for u in src)+"</ul>")
index(T,tabs,"Sections",cur="datactr");print(len(src),"sources")
