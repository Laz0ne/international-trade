const pptxgen = require("pptxgenjs");

const C = {
  navy:"1E2761", midnav:"162054", ice:"CADCFC", white:"FFFFFF",
  gold:"F0C040", slate:"8090B0", green:"34D399", red:"F87171",
  cardBg:"253180", lightBg:"F0F4FF", darkText:"1E2761",
};
const F = { title:"Trebuchet MS", body:"Calibri" };

function dark(p) { const s=p.addSlide(); s.background={color:C.navy}; return s; }
function light(p) { const s=p.addSlide(); s.background={color:C.lightBg}; return s; }
function badge(s, dk) {
  const bg=dk?C.cardBg:"E4ECFF", fc=dk?C.white:C.navy;
  s.addShape("rect",{x:8.8,y:0.12,w:1.05,h:0.32,fill:{color:bg},line:{color:bg,width:0}});
  s.addText("FERRARI & BORI",{x:8.8,y:0.12,w:1.05,h:0.32,fontSize:6,fontFace:F.body,bold:true,color:fc,align:"center",valign:"middle",margin:0});
}
function num(s,n,dk) { s.addText(n+" / 5",{x:0.3,y:5.3,w:0.5,h:0.2,fontSize:8,color:dk?C.slate:"9090B0",fontFace:F.body,align:"left"}); }
function bar(s,y,h,c) { s.addShape("rect",{x:0,y,w:0.09,h,fill:{color:c||C.gold},line:{color:c||C.gold,width:0}}); }
function mk() { return {type:"outer",blur:8,offset:3,angle:135,color:"000000",opacity:0.1}; }

async function build() {
  const pres = new pptxgen();
  pres.layout="LAYOUT_16x9";
  pres.title="Supply Chain as Growth Driver — CMO Presentation";

  // ── SLIDE 1: Executive Hook ──────────────────────────────────────────────
  {
    const s = dark(pres);
    badge(s,true);
    s.addShape("rect",{x:0,y:0,w:0.09,h:5.625,fill:{color:C.gold},line:{color:C.gold,width:0}});
    s.addText("EXECUTIVE SUMMARY",{x:0.35,y:0.48,w:9.0,h:0.26,fontSize:9,fontFace:F.body,color:C.gold,bold:true,charSpacing:4,align:"left",margin:0});
    s.addText("Our supply chain strategy guarantees\nfull product availability — at every demand peak.",{
      x:0.35,y:0.82,w:9.0,h:1.35,fontSize:22,fontFace:F.title,bold:true,
      color:C.white,align:"left",valign:"top",lineSpacingMultiple:1.2,margin:0});
    s.addText("Zero stockouts. Optimized costs. Revenue protected. 52 weeks.",{
      x:0.35,y:2.22,w:9.0,h:0.4,fontSize:13,fontFace:F.body,color:C.ice,italic:true,align:"left",margin:0});
    const stats=[{num:"100%",label:"Product\nAvailability"},{num:"0",label:"Stockout\nWeeks"},{num:"2×",label:"Air Shipments\nfor Peak Demand"}];
    stats.forEach((st,i)=>{
      const x=0.35+i*3.1;
      s.addShape("rect",{x,y:2.8,w:2.85,h:2.1,fill:{color:C.cardBg},line:{color:C.ice,width:1}});
      s.addShape("rect",{x,y:2.8,w:2.85,h:0.07,fill:{color:C.gold},line:{color:C.gold,width:0}});
      s.addText(st.num,{x,y:2.88,w:2.85,h:0.92,fontSize:38,fontFace:F.title,bold:true,color:C.gold,align:"center",margin:0});
      s.addText(st.label,{x:x+0.1,y:3.85,w:2.65,h:0.9,fontSize:12,fontFace:F.body,color:C.ice,align:"center",wrap:true,margin:0});
    });
    s.addText("Ferrari × Bori — International Product Launch | Supply Chain Strategy Presentation",{
      x:0.35,y:5.28,w:9.0,h:0.22,fontSize:8,fontFace:F.body,color:C.slate,italic:true,align:"left",margin:0});
    s.addText("1 / 5",{x:9.5,y:5.3,w:0.4,h:0.2,fontSize:8,color:C.slate,fontFace:F.body,align:"right"});
    s.addNotes("Open with conviction. We are not here to talk about shipping — we are here to talk about growth. Our strategy ensures that every product is available, every time. Zero stockouts, zero lost sales. That is the business result our supply chain was built to deliver.");
  }

  // ── SLIDE 2: The Problem ─────────────────────────────────────────────────
  {
    const s = light(pres);
    badge(s,false);
    num(s,2,false);
    bar(s,0,5.625,C.red);
    s.addText("THE PROBLEM WE SOLVED",{x:0.35,y:0.22,w:9.3,h:0.3,fontSize:9,fontFace:F.body,bold:true,color:C.red,charSpacing:3,align:"left",margin:0});
    s.addText("Before Our Fix: Revenue Was at Risk",{x:0.35,y:0.58,w:9.3,h:0.55,fontSize:25,fontFace:F.title,bold:true,color:C.darkText,align:"left",margin:0});
    const problems=[
      {icon:"⚠️",title:"Stockouts During\nPeak Weeks",body:"Weeks 33–34 = 6× normal demand with no coverage plan.\nResult: empty shelves, lost sales, frustrated customers."},
      {icon:"📦",title:"Inefficient Container\nLoading",body:"Poor packing wasted FEU space.\nResult: full shipping cost paid for under-loaded containers."},
      {icon:"📉",title:"No Inventory\nFlow Logic",body:"No week-by-week tracking from factory to warehouse.\nResult: unable to predict or prevent availability gaps."},
    ];
    problems.forEach((p,i)=>{
      const x=0.22+i*3.2;
      s.addShape("rect",{x,y:1.32,w:3.0,h:3.65,fill:{color:C.white},shadow:mk(),line:{color:"E0E8FF",width:1}});
      s.addShape("rect",{x,y:1.32,w:3.0,h:0.08,fill:{color:C.red},line:{color:C.red,width:0}});
      s.addText(p.icon,{x,y:1.5,w:3.0,h:0.55,fontSize:22,align:"center",margin:0});
      s.addText(p.title,{x,y:2.1,w:3.0,h:0.7,fontSize:13.5,fontFace:F.title,bold:true,color:C.darkText,align:"center",margin:0});
      s.addText(p.body,{x:x+0.15,y:2.88,w:2.7,h:1.85,fontSize:11,fontFace:F.body,color:"445580",align:"left",valign:"top",wrap:true,margin:0});
    });
    s.addShape("rect",{x:0.22,y:5.06,w:9.55,h:0.35,fill:{color:"FFF0F0"},line:{color:C.red,width:1}});
    s.addText("📌  Business impact: Unreliable availability = customer dissatisfaction = brand erosion = lost market share",{
      x:0.35,y:5.07,w:9.3,h:0.32,fontSize:10,fontFace:F.body,bold:true,color:C.red,align:"left",valign:"middle",margin:0});
    s.addNotes("Before our corrected model, we had a plan that could not survive real demand. The biggest risk was weeks 33 and 34 — six times normal volume — with no air backup. That is not a logistics problem. That is lost revenue and damaged customer trust.");
  }

  // ── SLIDE 3: Our Strategy ────────────────────────────────────────────────
  {
    const s = dark(pres);
    badge(s,true);
    num(s,3,true);
    s.addShape("rect",{x:0,y:0,w:0.09,h:5.625,fill:{color:C.green},line:{color:C.green,width:0}});
    s.addText("OUR SOLUTION",{x:0.35,y:0.22,w:9.3,h:0.28,fontSize:9,fontFace:F.body,bold:true,color:C.green,charSpacing:3,align:"left",margin:0});
    s.addText("A Three-Layer Strategy Built Around Customer Demand",{x:0.35,y:0.58,w:9.3,h:0.6,fontSize:21,fontFace:F.title,bold:true,color:C.white,align:"left",margin:0});
    const pillars=[
      {num:"01",title:"Ocean Freight\n(Cost Efficiency)",bullets:["FEU containers fully optimized","Maximized packing = lower cost per unit","Scheduled weekly to maintain base stock","→ Enables competitive pricing"],color:C.ice},
      {num:"02",title:"Air Shipments\n(Demand Protection)",bullets:["2 targeted air shipments at peak weeks","Weeks 33–34: 6× surge demand covered","4-week lead time vs 10 by ocean","→ Product on shelf when demand spikes"],color:C.gold},
      {num:"03",title:"Weekly Inventory\nTracking",bullets:["Week-by-week flow: factory → warehouse","Zero stockout weeks guaranteed","Buffer stock built before every peak","→ Customers never face an empty shelf"],color:C.green},
    ];
    pillars.forEach((p,i)=>{
      const x=0.22+i*3.22;
      s.addShape("rect",{x,y:1.35,w:3.0,h:3.9,fill:{color:C.cardBg},line:{color:p.color,width:1}});
      s.addShape("rect",{x,y:1.35,w:3.0,h:0.07,fill:{color:p.color},line:{color:p.color,width:0}});
      s.addText(p.num,{x,y:1.42,w:3.0,h:0.68,fontSize:34,fontFace:F.title,bold:true,color:p.color,align:"center",margin:0});
      s.addText(p.title,{x,y:2.12,w:3.0,h:0.68,fontSize:13,fontFace:F.title,bold:true,color:C.white,align:"center",margin:0});
      const items=p.bullets.map((b,bi)=>({
        text:b, options:{
          bullet: bi<p.bullets.length-1 ? true : false,
          bold: bi===p.bullets.length-1,
          color: bi===p.bullets.length-1 ? p.color : C.ice,
          breakLine: bi<p.bullets.length-1
        }
      }));
      s.addText(items,{x:x+0.18,y:2.88,w:2.65,h:2.25,fontSize:10.5,fontFace:F.body,valign:"top",align:"left",margin:0});
    });
    s.addNotes("We built a three-layer system: cost-efficient ocean freight for the baseline, rapid air shipments to protect our two biggest demand weeks, and a week-by-week inventory model so we can see exactly when stock arrives and when it is needed. Every layer exists for one reason — to make sure the product is there when the customer wants it.");
  }

  // ── SLIDE 4: Business Impact ─────────────────────────────────────────────
  {
    const s = light(pres);
    badge(s,false);
    num(s,4,false);
    bar(s,0,5.625,C.gold);
    s.addText("BUSINESS IMPACT",{x:0.35,y:0.22,w:9.3,h:0.28,fontSize:9,fontFace:F.body,bold:true,color:"C0900A",charSpacing:3,align:"left",margin:0});
    s.addText("From Operational Fix to Revenue Driver",{x:0.35,y:0.58,w:7.0,h:0.55,fontSize:23,fontFace:F.title,bold:true,color:C.darkText,align:"left",margin:0});
    const impacts=[
      {icon:"🛡️",title:"Revenue Protection",text:"Zero stockout weeks = zero lost sales.\nEvery demand week is covered — especially the critical 6× peak.\nRevenue is never left on the table.",color:C.navy},
      {icon:"😊",title:"Customer Experience",text:"Product always available when customers look for it.\nNo backorders, no substitute products.\nBuilds loyalty and repeat purchase intent.",color:"2563EB"},
      {icon:"⚡",title:"Speed to Market",text:"Air shipments cut lead time from 10 weeks to 4.\nWe respond to unexpected demand 60% faster.\nFirst-mover advantage maintained at peak periods.",color:"7C3AED"},
      {icon:"💰",title:"Optimized Cost Structure",text:"Ocean freight handles 95% of volume at low cost.\nAir used surgically — only when revenue risk justifies it.\nCost efficiency creates budget for marketing activation.",color:"059669"},
    ];
    impacts.forEach((imp,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const x=0.22+col*4.75, y=1.32+row*2.08;
      s.addShape("rect",{x,y,w:4.5,h:1.88,fill:{color:C.white},shadow:mk(),line:{color:"E0E8FF",width:1}});
      s.addShape("rect",{x,y,w:0.07,h:1.88,fill:{color:imp.color},line:{color:imp.color,width:0}});
      s.addText(imp.icon+"  "+imp.title,{x:x+0.2,y:y+0.12,w:4.1,h:0.36,fontSize:13,fontFace:F.title,bold:true,color:imp.color,align:"left",margin:0});
      s.addText(imp.text,{x:x+0.2,y:y+0.52,w:4.1,h:1.22,fontSize:10.5,fontFace:F.body,color:"445580",align:"left",valign:"top",wrap:true,margin:0});
    });
    s.addNotes("Let me be direct about what this means for the business. Every week without a stockout is revenue that stays in our pocket. The air shipment cost is small compared to the revenue we would lose during a 6× demand week with empty shelves. And the savings from ocean freight give us flexibility to invest in marketing.");
  }

  // ── SLIDE 5: Recommendation ──────────────────────────────────────────────
  {
    const s = dark(pres);
    badge(s,true);
    num(s,5,true);
    s.addShape("rect",{x:0,y:0,w:10,h:0.08,fill:{color:C.gold},line:{color:C.gold,width:0}});
    s.addText("RECOMMENDATION",{x:0.5,y:0.22,w:9.0,h:0.28,fontSize:9,fontFace:F.body,bold:true,color:C.gold,charSpacing:3,align:"left",margin:0});
    s.addText("Approve This Strategy — Supply Chain is Our Growth Engine",{x:0.5,y:0.6,w:9.0,h:0.6,fontSize:20,fontFace:F.title,bold:true,color:C.white,align:"left",margin:0});

    s.addShape("rect",{x:0.3,y:1.38,w:4.6,h:2.8,fill:{color:C.cardBg},line:{color:C.ice,width:1}});
    s.addText("WHAT WE RECOMMEND",{x:0.5,y:1.48,w:4.2,h:0.28,fontSize:8.5,fontFace:F.body,bold:true,color:C.gold,charSpacing:2,align:"left",margin:0});
    const recs=[
      "Adopt the corrected 52-week inventory plan",
      "Confirm 2 air shipments for peak demand weeks",
      "Use fully optimized FEU containers for all ocean freight",
      "Review inventory model quarterly vs. actual demand",
    ];
    const recItems=recs.map((r,i)=>({text:r,options:{bullet:true,color:C.ice,breakLine:i<recs.length-1}}));
    s.addText(recItems,{x:0.5,y:1.85,w:4.15,h:2.1,fontSize:11,fontFace:F.body,valign:"top",margin:0});

    s.addShape("rect",{x:5.1,y:1.38,w:4.6,h:2.8,fill:{color:C.cardBg},line:{color:C.gold,width:1}});
    s.addText("MARKETING KPI IMPACT",{x:5.3,y:1.48,w:4.2,h:0.28,fontSize:8.5,fontFace:F.body,bold:true,color:C.gold,charSpacing:2,align:"left",margin:0});
    const kpis=[
      {label:"Sales Continuity",val:"52 / 52 weeks covered"},
      {label:"Stockout Risk",val:"Eliminated"},
      {label:"Brand Reliability",val:"Product always available"},
      {label:"Customer Sat.",val:"No lost purchase intent"},
      {label:"Peak Readiness",val:"6× surge absorbed"},
    ];
    kpis.forEach((k,i)=>{
      const y=1.9+i*0.44;
      s.addText(k.label,{x:5.3,y,w:2.3,h:0.35,fontSize:10.5,fontFace:F.body,color:C.slate,align:"left",valign:"middle",margin:0});
      s.addText(k.val,{x:7.65,y,w:1.8,h:0.35,fontSize:10.5,fontFace:F.body,bold:true,color:C.green,align:"right",valign:"middle",margin:0});
      if(i<kpis.length-1) s.addShape("line",{x:5.3,y:y+0.36,w:4.1,h:0,line:{color:"303E70",width:0.5}});
    });

    s.addShape("rect",{x:0.3,y:4.38,w:9.4,h:0.78,fill:{color:C.gold},line:{color:C.gold,width:0}});
    s.addText("Supply chain is not a cost center — it is the foundation that makes every marketing promise deliverable.",{
      x:0.5,y:4.4,w:9.1,h:0.72,fontSize:13,fontFace:F.title,bold:true,color:C.navy,align:"center",valign:"middle",margin:0});
    s.addNotes("We are asking for your approval to implement this plan in full. Every promise marketing makes to the customer — that the product is available, on time, every time — is only as strong as the supply chain behind it. We have built that supply chain. The question is: are we ready to grow?");
  }

  await pres.writeFile({ fileName: "CMO_Deck_Ferrari_Bori_FINAL.pptx" });
  console.log("✅ Done — CMO_Deck_Ferrari_Bori_FINAL.pptx generated!");
}
build().catch(console.error);