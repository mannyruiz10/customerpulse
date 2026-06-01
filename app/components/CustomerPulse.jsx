import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

const FONTS = "https://fonts.googleapis.com/css2?family=Cabinet+Grotesk:wght@400;500;700;800;900&family=Satoshi:wght@300;400;500;700&display=swap";

const CSS = `
@import url('${FONTS}');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root {
  --navy:#050810;--navy2:#0a0e1a;
  --indigo:#5B4FFF;--indigo2:#7B70FF;--indigo3:rgba(91,79,255,0.12);
  --amber:#FF9F1C;--amber2:#FFB547;--amber3:rgba(255,159,28,0.12);
  --coral:#FF4D6D;--coral3:rgba(255,77,109,0.12);
  --emerald:#0BC47F;--emerald3:rgba(11,196,127,0.12);
  --text:#D4C5A9;--text2:#A89880;--text3:#6B5E4E;
  --border:rgba(255,255,255,0.06);--border2:rgba(255,255,255,0.1);
  --glass:rgba(255,255,255,0.03);--glass2:rgba(255,255,255,0.06);
  --fd:'Cabinet Grotesk',sans-serif;--fb:'Satoshi',sans-serif;
  --r:10px;--r2:16px;--r3:24px;
}
html{scroll-behavior:smooth}
body{background:var(--navy);color:var(--text);font-family:var(--fb);min-height:100vh;overflow-x:hidden}
h1,h2,h3,h4{font-family:var(--fd)}

@keyframes fadeUp{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:translateY(0)}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes shimmer{0%{background-position:-200% center}100%{background-position:200% center}}
@keyframes pulse-glow{0%,100%{box-shadow:0 0 0 0 rgba(255,159,28,0)}50%{box-shadow:0 0 0 8px rgba(255,159,28,0.15)}}
@keyframes border-dance{0%,100%{border-color:rgba(91,79,255,0.3)}50%{border-color:rgba(91,79,255,0.7);box-shadow:0 0 30px rgba(91,79,255,0.15)}}
@keyframes bar-grow{from{width:0}to{width:var(--w)}}
@keyframes load-bar{0%{width:0%}30%{width:40%}70%{width:75%}100%{width:95%}}
@keyframes count{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}

.fu{animation:fadeUp 0.55s cubic-bezier(0.22,1,0.36,1) both}
.fu1{animation-delay:0.08s}.fu2{animation-delay:0.16s}.fu3{animation-delay:0.24s}
.fu4{animation-delay:0.32s}.fu5{animation-delay:0.40s}.fu6{animation-delay:0.48s}

.shimmer-text{
  background:linear-gradient(90deg,var(--text) 0%,var(--amber2) 40%,var(--text) 80%,var(--indigo2) 100%);
  background-size:200% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;
  background-clip:text;animation:shimmer 4s linear infinite
}
.gradient-text{
  background:linear-gradient(135deg,var(--indigo2),var(--amber));
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text
}

.btn-cta{
  display:inline-flex;align-items:center;gap:10px;
  background:linear-gradient(135deg,var(--amber) 0%,#FF6B00 100%);
  color:var(--navy);font-family:var(--fd);font-weight:800;font-size:1rem;
  padding:16px 36px;border-radius:100px;border:none;cursor:pointer;transition:all 0.25s;
  box-shadow:0 4px 24px rgba(255,159,28,0.35),0 1px 0 rgba(255,255,255,0.2) inset;
  animation:pulse-glow 3s ease infinite;letter-spacing:-0.01em
}
.btn-cta:hover{transform:translateY(-2px) scale(1.02);box-shadow:0 8px 32px rgba(255,159,28,0.45)}
.btn-cta:disabled{opacity:0.45;cursor:not-allowed;transform:none;animation:none;box-shadow:none}

.btn-secondary{
  display:inline-flex;align-items:center;gap:8px;background:var(--glass2);color:var(--text);
  font-family:var(--fd);font-weight:700;font-size:0.9rem;padding:14px 28px;border-radius:100px;
  border:1px solid var(--border2);cursor:pointer;transition:all 0.2s;backdrop-filter:blur(10px)
}
.btn-secondary:hover{background:var(--glass);border-color:rgba(255,255,255,0.2);transform:translateY(-1px)}

.btn-ghost{
  display:inline-flex;align-items:center;gap:6px;background:transparent;color:var(--text2);
  font-family:var(--fb);font-size:0.875rem;padding:8px 16px;border-radius:var(--r);
  border:1px solid var(--border);cursor:pointer;transition:all 0.2s
}
.btn-ghost:hover{color:var(--text);border-color:var(--border2);background:var(--glass)}

.btn-plan{
  display:inline-flex;align-items:center;justify-content:center;gap:8px;width:100%;
  font-family:var(--fd);font-weight:700;font-size:0.9rem;padding:14px 24px;
  border-radius:100px;border:1px solid var(--border2);cursor:pointer;transition:all 0.2s;
  background:transparent;color:var(--text)
}
.btn-plan:hover{border-color:var(--indigo);background:var(--indigo3);color:var(--indigo2)}
.btn-plan-cta{
  display:inline-flex;align-items:center;justify-content:center;gap:8px;width:100%;
  font-family:var(--fd);font-weight:800;font-size:0.9rem;padding:14px 24px;
  border-radius:100px;border:none;cursor:pointer;transition:all 0.25s;
  background:linear-gradient(135deg,var(--amber),#FF6B00);color:var(--navy);
  box-shadow:0 4px 20px rgba(255,159,28,0.3)
}
.btn-plan-cta:hover{transform:translateY(-1px);box-shadow:0 6px 28px rgba(255,159,28,0.4)}

.card{
  background:var(--glass);border:1px solid var(--border);border-radius:var(--r3);
  backdrop-filter:blur(20px);position:relative;overflow:hidden
}
.card::before{
  content:'';position:absolute;inset:0;border-radius:inherit;
  background:linear-gradient(135deg,rgba(255,255,255,0.04) 0%,transparent 60%);pointer-events:none
}
.card-glow{transition:border-color 0.3s,box-shadow 0.3s}
.card-glow:hover{border-color:rgba(91,79,255,0.3);box-shadow:0 0 40px rgba(91,79,255,0.08)}
.card-coral{border-color:rgba(255,77,109,0.2);background:rgba(255,77,109,0.04)}
.card-emerald{border-color:rgba(11,196,127,0.2);background:rgba(11,196,127,0.04)}
.card-indigo{border-color:rgba(91,79,255,0.2);background:rgba(91,79,255,0.04)}
.card-featured{border-color:rgba(255,159,28,0.35);background:rgba(255,159,28,0.04)}

.badge{
  display:inline-flex;align-items:center;gap:5px;font-size:0.72rem;font-weight:700;
  letter-spacing:0.07em;text-transform:uppercase;padding:5px 12px;border-radius:100px;font-family:var(--fd)
}
.badge-amber{background:var(--amber3);color:var(--amber2);border:1px solid rgba(255,159,28,0.2)}
.badge-indigo{background:var(--indigo3);color:var(--indigo2);border:1px solid rgba(91,79,255,0.2)}
.badge-emerald{background:var(--emerald3);color:var(--emerald);border:1px solid rgba(11,196,127,0.2)}
.badge-coral{background:var(--coral3);color:var(--coral);border:1px solid rgba(255,77,109,0.2)}

select,textarea{
  background:rgba(255,255,255,0.04);border:1px solid var(--border2);color:var(--text);
  border-radius:var(--r2);font-family:var(--fb);font-size:0.9rem;outline:none;transition:all 0.2s;width:100%
}
select{padding:14px 18px;cursor:pointer;appearance:none}
select:focus,textarea:focus{border-color:var(--indigo);box-shadow:0 0 0 3px rgba(91,79,255,0.12)}
textarea{padding:16px 18px;resize:vertical;min-height:160px;line-height:1.65}
select option{background:var(--navy2)}

.upload-zone{
  border:2px dashed rgba(91,79,255,0.35);border-radius:var(--r3);padding:52px 28px;
  text-align:center;cursor:pointer;transition:all 0.3s;background:rgba(91,79,255,0.03);
  position:relative;overflow:hidden
}
.upload-zone:hover,.upload-zone.drag-over{border-color:var(--indigo);background:rgba(91,79,255,0.06);box-shadow:0 0 40px rgba(91,79,255,0.1)}
.upload-zone.idle{animation:border-dance 4s ease infinite}

.prog-track{height:8px;background:rgba(255,255,255,0.05);border-radius:4px;overflow:hidden}
.prog-fill{height:100%;border-radius:4px}

.step-dot{
  width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-family:var(--fd);font-size:0.8rem;font-weight:800;border:1.5px solid var(--border2);
  color:var(--text3);background:var(--glass);transition:all 0.3s;flex-shrink:0
}
.step-dot.active{border-color:var(--indigo);color:var(--indigo2);background:var(--indigo3)}
.step-dot.done{border-color:var(--emerald);color:var(--navy);background:var(--emerald)}

.im-tab{
  display:flex;flex-direction:column;align-items:center;gap:8px;padding:18px 14px;
  border-radius:var(--r2);border:1px solid var(--border2);background:transparent;
  cursor:pointer;transition:all 0.2s;flex:1
}
.im-tab:hover{border-color:var(--indigo2);background:var(--indigo3)}
.im-tab.active{border-color:var(--indigo);background:rgba(91,79,255,0.1);box-shadow:0 0 20px rgba(91,79,255,0.1)}
.im-tab .im-emoji{font-size:1.4rem;line-height:1}
.im-tab .im-label{font-size:0.78rem;font-weight:600;font-family:var(--fd);color:var(--text2);text-align:center}
.im-tab.active .im-label{color:var(--indigo2)}

.sec-head{display:flex;align-items:center;gap:10px;margin-bottom:20px}
.sec-icon{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0}
.sec-title{font-family:var(--fd);font-weight:700;font-size:0.9rem;letter-spacing:0.02em}
.sec-letter{font-size:0.65rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:var(--text3);padding:2px 8px;background:var(--glass2);border-radius:100px;border:1px solid var(--border)}

.toggle-track{width:44px;height:24px;border-radius:12px;border:none;cursor:pointer;position:relative;transition:background 0.25s;flex-shrink:0}
.toggle-thumb{position:absolute;top:3px;width:18px;height:18px;border-radius:50%;background:white;transition:left 0.25s cubic-bezier(0.34,1.56,0.64,1);box-shadow:0 2px 4px rgba(0,0,0,0.3)}

.load-bar{height:3px;background:linear-gradient(90deg,var(--indigo),var(--amber));animation:load-bar 3s ease-out forwards;border-radius:2px}
.spinner{width:44px;height:44px;border:3px solid rgba(91,79,255,0.2);border-top-color:var(--indigo);border-radius:50%;animation:spin 0.75s linear infinite}

.report-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}
@media(max-width:700px){.report-grid{grid-template-columns:1fr}}

.col-sel-btn{
  padding:8px 16px;border-radius:var(--r);font-size:0.8rem;font-weight:600;font-family:var(--fd);
  border:1px solid var(--border2);background:transparent;color:var(--text2);cursor:pointer;transition:all 0.2s
}
.col-sel-btn:hover,.col-sel-btn.active{border-color:var(--indigo);background:var(--indigo3);color:var(--indigo2)}

.error-box{padding:14px 18px;border-radius:var(--r2);background:var(--coral3);border:1px solid rgba(255,77,109,0.25);color:var(--coral);font-size:0.875rem;display:flex;gap:10px;align-items:flex-start}
.tag{font-size:0.75rem;padding:4px 12px;border-radius:100px;background:var(--glass2);border:1px solid var(--border2);color:var(--text2);font-family:var(--fd);font-weight:600}

.check-row{display:flex;align-items:flex-start;gap:10px;margin-bottom:12px}
.check-icon{width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:0.65rem;flex-shrink:0;margin-top:1px}
.check-yes{background:var(--emerald3);border:1px solid rgba(11,196,127,0.3);color:var(--emerald)}
.check-no{background:rgba(255,255,255,0.04);border:1px solid var(--border);color:var(--text3)}

.comp-table{width:100%;border-collapse:collapse}
.comp-table th,.comp-table td{padding:14px 18px;text-align:center;border-bottom:1px solid var(--border);font-size:0.875rem;color:var(--text2)}
.comp-table th{font-family:var(--fd);font-weight:700;font-size:0.8rem;letter-spacing:0.04em;color:var(--text);text-transform:uppercase}
.comp-table td:first-child{text-align:left;color:var(--text)}
.comp-table tr:last-child td{border-bottom:none}
.comp-col-us{background:rgba(91,79,255,0.06);color:var(--indigo2)!important;font-weight:700!important;font-family:var(--fd)!important}
.comp-head-us{background:rgba(91,79,255,0.1);color:var(--indigo2)!important;border-radius:var(--r) var(--r) 0 0}

.nav-link{background:none;border:none;color:var(--text2);font-family:var(--fb);font-size:0.875rem;cursor:pointer;padding:6px 12px;border-radius:var(--r);transition:color 0.2s}
.nav-link:hover{color:var(--text)}

section{padding:100px 24px}
.section-inner{max-width:1100px;margin:0 auto}
.section-label{font-size:0.72rem;font-weight:800;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2);font-family:var(--fd);margin-bottom:16px}
.section-title{font-size:clamp(1.8rem,4vw,2.8rem);font-weight:900;line-height:1.15;letter-spacing:-0.02em;margin-bottom:20px}
.section-sub{font-size:1.05rem;color:var(--text2);line-height:1.75;max-width:560px}
p{color:inherit}

.divider{height:1px;background:var(--border);max-width:1100px;margin:0 auto}

@media(max-width:768px){
  section{padding:70px 20px}
  .hide-mobile{display:none!important}
}
`;

function mockAnalyze(reviews, businessType) {
  const text=reviews.join(" ").toLowerCase();
  const pos=["great","excellent","love","amazing","perfect","good","fantastic","wonderful","best","happy","satisfied","recommend","fast","friendly","clean","fresh","quality"];
  const neg=["bad","terrible","awful","poor","worst","slow","rude","cold","wrong","broken","late","dirty","expensive","disappointed","never","horrible","disgusting","wait","long","small"];
  let p=0,n=0;
  pos.forEach(w=>{const m=text.match(new RegExp(`\\b${w}\\b`,"g"));if(m)p+=m.length});
  neg.forEach(w=>{const m=text.match(new RegExp(`\\b${w}\\b`,"g"));if(m)n+=m.length});
  const score=Math.max(18,Math.min(94,Math.round((p/(p+n||1))*100)));
  const sentiment=score>=70?"Mostly Positive":score>=45?"Mixed":"Mostly Negative";
  const T={
    Restaurant:{
      complaints:[{theme:"Slow Service",pct:34,color:"#FF4D6D"},{theme:"Cold / Wrong Food",pct:27,color:"#FF9F1C"},{theme:"Long Wait Times",pct:22,color:"#FBBF24"},{theme:"Staff Attitude",pct:11,color:"#F97316"},{theme:"Cleanliness",pct:6,color:"#EF4444"}],
      positives:[{theme:"Food Quality & Taste",pct:41,color:"#0BC47F"},{theme:"Friendly Staff",pct:26,color:"#06B6D4"},{theme:"Great Atmosphere",pct:18,color:"#5B4FFF"},{theme:"Good Value",pct:15,color:"#8B5CF6"}],
      requests:["More vegetarian / vegan options","Faster service during peak hours","Online reservation system","Better parking","Loyalty rewards program"],
      actions:[{emoji:"⚡",title:"Hire peak-hour support staff",detail:"Add 2 part-time staff for Friday–Sunday dinner service. Service speed is your #1 complaint and the fastest win available."},{emoji:"🌡️",title:"Implement food temperature checkpoints",detail:"Add a plating quality check before every dish leaves the kitchen. Train staff on holding temperatures."},{emoji:"💬",title:"Respond to every review within 48h",detail:"Assign someone to reply to all Google/Yelp reviews. Responding to negatives reduces churn by ~30%."},{emoji:"🌿",title:"Expand vegetarian & vegan menu",detail:"Add 3–4 plant-based dishes. Customers are asking — it broadens your market significantly."},{emoji:"🎁",title:"Launch a digital loyalty program",detail:"A simple stamp card via Square or Toast increases repeat visits by an average of 21%."}],
      opportunity:"Fixing service speed alone could recover 30–40% of lost customers. Every slow-service complaint is a table that may never return. Staffing up during peak hours has the fastest ROI of any change you can make today.",
    },
    "Retail Store":{
      complaints:[{theme:"Out of Stock Items",pct:38,color:"#FF4D6D"},{theme:"Long Checkout Lines",pct:29,color:"#FF9F1C"},{theme:"Unhelpful Staff",pct:18,color:"#FBBF24"},{theme:"Poor Store Layout",pct:10,color:"#F97316"},{theme:"Return Policy",pct:5,color:"#EF4444"}],
      positives:[{theme:"Product Selection",pct:38,color:"#0BC47F"},{theme:"Clean Store",pct:28,color:"#06B6D4"},{theme:"Convenient Location",pct:21,color:"#5B4FFF"},{theme:"Competitive Pricing",pct:13,color:"#8B5CF6"}],
      requests:["More size/color variety","Self-checkout lanes","Online inventory visibility","Easier returns process","Staff product knowledge"],
      actions:[{emoji:"📦",title:"Improve real-time inventory tracking",detail:"Implement alerts so staff restock before items run out."},{emoji:"💳",title:"Add self-checkout options",detail:"Even 1–2 kiosks can cut wait times 40% and improve satisfaction significantly."},{emoji:"📚",title:"Weekly staff product training",detail:"15-minute briefings on new arrivals turn staff into trusted advisors."},{emoji:"🔄",title:"Simplify your return policy",detail:"A clear 30-day return policy is baseline today — communicate it prominently."},{emoji:"📱",title:"Launch restock alert notifications",detail:"Let customers sign up for restock alerts to capture every lost sale."}],
      opportunity:"Stockout issues are silently costing you sales every day. Better inventory visibility could recapture 15–25% of walk-away revenue.",
    },
    "Ecommerce Product":{
      complaints:[{theme:"Slow / Late Shipping",pct:41,color:"#FF4D6D"},{theme:"Product Quality Issues",pct:28,color:"#FF9F1C"},{theme:"Poor Packaging",pct:16,color:"#FBBF24"},{theme:"Confusing Sizing",pct:10,color:"#F97316"},{theme:"Customer Support",pct:5,color:"#EF4444"}],
      positives:[{theme:"Value for Money",pct:37,color:"#0BC47F"},{theme:"Accurate Product Photos",pct:24,color:"#06B6D4"},{theme:"Easy Ordering",pct:22,color:"#5B4FFF"},{theme:"Fast Delivery (when on time)",pct:17,color:"#8B5CF6"}],
      requests:["More size options","Faster shipping tier","Better size chart","Bundle deals","Subscription option"],
      actions:[{emoji:"🚚",title:"Audit your fulfillment partner",detail:"Late shipping is your #1 complaint. Benchmark your carrier's on-time rate."},{emoji:"📐",title:"Add a detailed size & fit guide",detail:"A proper fit guide with measurements reduces return rate by up to 23%."},{emoji:"📦",title:"Upgrade packaging quality",detail:"Unboxing is a shareable moment that drives organic social content."},{emoji:"⭐",title:"Automate post-delivery review requests",detail:"Send an email 7 days after delivery. Happy customers won't review without being asked."},{emoji:"💰",title:"Test a bundle or subscription offer",detail:"A simple 'Save 15% on 3+' bundle can lift average order value 20–35%."}],
      opportunity:"Shipping speed is your biggest revenue risk. A premium shipping tier could improve satisfaction and generate extra margin simultaneously.",
    },
    "Service Business":{
      complaints:[{theme:"Response Time",pct:36,color:"#FF4D6D"},{theme:"Communication",pct:27,color:"#FF9F1C"},{theme:"Pricing Transparency",pct:20,color:"#FBBF24"},{theme:"Quality Consistency",pct:12,color:"#F97316"},{theme:"Scheduling",pct:5,color:"#EF4444"}],
      positives:[{theme:"Knowledgeable Staff",pct:40,color:"#0BC47F"},{theme:"Professional Results",pct:28,color:"#06B6D4"},{theme:"Convenient Hours",pct:18,color:"#5B4FFF"},{theme:"Clean Environment",pct:14,color:"#8B5CF6"}],
      requests:["Online booking","Upfront pricing / estimates","Status updates","After-hours availability","Membership pricing"],
      actions:[{emoji:"📅",title:"Implement online booking",detail:"Tools like Calendly eliminate phone tag and reduce no-shows by 29%."},{emoji:"💬",title:"Send automated status updates",detail:"A simple SMS at job start and completion eliminates customer anxiety."},{emoji:"💵",title:"Publish transparent pricing tiers",detail:"Surprise pricing is the #1 driver of bad reviews."},{emoji:"⭐",title:"Create a review follow-up system",detail:"Send a review request 24h after service completion."},{emoji:"📦",title:"Introduce service packages",detail:"Monthly packages improve cash flow and lock in loyalty."}],
      opportunity:"Response time and communication are almost free to fix and they're your top complaints. Automated updates could transform your worst reviews into your best ones within 30 days.",
    },
    Other:{
      complaints:[{theme:"Value / Pricing",pct:33,color:"#FF4D6D"},{theme:"Service Quality",pct:27,color:"#FF9F1C"},{theme:"Communication",pct:22,color:"#FBBF24"},{theme:"Consistency",pct:11,color:"#F97316"},{theme:"Availability",pct:7,color:"#EF4444"}],
      positives:[{theme:"Quality of Offering",pct:40,color:"#0BC47F"},{theme:"Friendly Team",pct:27,color:"#06B6D4"},{theme:"Good Location",pct:19,color:"#5B4FFF"},{theme:"Reliable",pct:14,color:"#8B5CF6"}],
      requests:["More variety","Better communication","Faster turnaround","Loyalty rewards","Online booking"],
      actions:[{emoji:"💬",title:"Respond to all reviews publicly",detail:"Responding signals professionalism and boosts local search rankings."},{emoji:"📊",title:"Send a monthly feedback survey",detail:"A simple Google Form catches issues before they become public."},{emoji:"🎯",title:"Clarify your value communication",detail:"Pricing complaints often mean the value isn't obvious. Update your messaging."},{emoji:"🔁",title:"Build a referral program",detail:"A 'get $10 for a referral' program can double word-of-mouth growth."},{emoji:"📱",title:"Complete your Google Business Profile",detail:"Add hours, photos, and regular posts. Free and high-impact."}],
      opportunity:"Pricing perception is usually a communication problem. Better storytelling on your site and in follow-ups can shift sentiment fast.",
    },
  };
  const t=T[businessType]||T["Other"];
  return {
    summary:`Analysis of ${reviews.length} customer reviews reveals ${sentiment.toLowerCase()} sentiment with a score of ${score}/100. ${score>=70?`Customers consistently praise your core offering — a strong foundation to build on. Targeted improvements can deliver quick wins.`:score>=45?`You have a loyal base of happy customers, but a meaningful segment experiences recurring friction. Addressing the top complaints could shift sentiment significantly.`:`Consistent pain points are affecting satisfaction. Most complaints cluster around a few fixable issues — focused action can lead to rapid improvement.`} The data points to ${t.complaints[0].theme.toLowerCase()} as the single highest-leverage improvement area right now.`,
    complaints:t.complaints,positives:t.positives,requests:t.requests,
    actions:t.actions,opportunity:t.opportunity,score,sentiment,reviewCount:reviews.length,
  };
}

async function aiAnalyze(reviews,businessType) {
  const sample=reviews.slice(0,80).join("\n---\n");
  const prompt=`You are a senior business analyst. Analyze these ${reviews.length} customer reviews for a ${businessType} and return ONLY valid JSON (no markdown) with this structure:
{"summary":"string","complaints":[{"theme":"string","pct":number,"color":"hex"}],"positives":[{"theme":"string","pct":number,"color":"hex"}],"requests":["string"],"actions":[{"emoji":"string","title":"string","detail":"string"}],"opportunity":"string","score":number,"sentiment":"Mostly Positive|Mixed|Mostly Negative","reviewCount":${reviews.length}}
4-6 items each for complaints/positives, 5 requests, 5 actions, score 1-100.
Complaint colors:#FF4D6D #FF9F1C #FBBF24 #F97316 #EF4444
Positive colors:#0BC47F #06B6D4 #5B4FFF #8B5CF6 #1DD1A1
REVIEWS:\n${sample}`;
  const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:prompt}]})});
  const data=await res.json();
  const text=data.content.map(b=>b.text||"").join("");
  return JSON.parse(text.replace(/```json|```/g,"").trim());
}

const SAMPLES={
  Restaurant:["The food was absolutely incredible! Best pasta I've had in years. Service was a bit slow though.","Love this place. Staff are so friendly and the atmosphere is perfect for a date night.","Waited 45 minutes for our table even with a reservation. Burger was cold when it arrived.","Great food, great prices. The salmon was perfect. Definitely coming back!","Terrible experience. Waitress forgot our drinks and the manager didn't care.","Amazing pizza — authentic and fresh. Wish they had more vegetarian options.","Good food but service ruined it. Ignored for 15 minutes before anyone took our order.","One of the best restaurants in town. Desserts are divine and the wine list is excellent.","The place looks nice but the food quality has declined. Used to be my favorite.","Quick service, delicious food, very clean. Exactly what I wanted!","Pasta was overcooked and the sauce was bland. For these prices I expected better.","Fantastic brunch spot! Eggs benedict were perfect. Only complaint is parking."],
  "Ecommerce Product":["Product arrived 2 weeks late and the packaging was crushed. Item seems fine but disappointed.","Great quality for the price! Exactly as described. Fast delivery too. Will order again.","Sizing runs small — ordered a medium and it fits like a small. Need a better size chart.","Love the product! Material is high quality and looks even better in person.","This is my third order. Consistent quality and usually arrives on time. Highly recommend.","Shipping took 3 weeks and customer service was unresponsive when I asked about my order.","The product is okay but not worth the price. Seen similar items for much less elsewhere.","Perfect! Exactly what I needed. Quick shipping and the packaging was nice and secure.","Bought two, one had a defect. Return process was easy at least, but frustrating overall.","Good product, wish they offered more color options. Would buy again if they expand the range."],
};

function Orbs() {
  return (
    <>
      <div style={{position:"fixed",inset:0,zIndex:0,pointerEvents:"none",background:"radial-gradient(ellipse 70% 50% at 15% -5%, rgba(91,79,255,0.22) 0%, transparent 55%), radial-gradient(ellipse 50% 40% at 85% 105%, rgba(255,159,28,0.14) 0%, transparent 50%)"}}/>
      <div style={{position:"fixed",inset:0,zIndex:0,pointerEvents:"none",backgroundImage:"linear-gradient(rgba(91,79,255,0.035) 1px, transparent 1px), linear-gradient(90deg, rgba(91,79,255,0.035) 1px, transparent 1px)",backgroundSize:"60px 60px",maskImage:"radial-gradient(ellipse 100% 80% at 50% 0%, black 20%, transparent 75%)"}}/>
    </>
  );
}

function Nav({onLogoClick, onLaunch, appMode}) {
  return (
    <nav style={{position:"sticky",top:0,zIndex:10,padding:"16px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",borderBottom:"1px solid var(--border)",backdropFilter:"blur(20px)",background:"rgba(10,15,30,0.85)"}}>
      <button onClick={onLogoClick} style={{display:"flex",alignItems:"center",gap:10,background:"none",border:"none",cursor:"pointer",color:"var(--text)"}}>
        <div style={{width:32,height:32,borderRadius:9,background:"linear-gradient(135deg,var(--indigo),var(--indigo2))",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1rem",boxShadow:"0 4px 12px rgba(91,79,255,0.4)"}}>📊</div>
        <span style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"1.1rem",letterSpacing:"-0.02em"}}>CustomerPulse</span>
      </button>
      {!appMode && (
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button className="nav-link hide-mobile" onClick={()=>document.getElementById("how")?.scrollIntoView({behavior:"smooth"})}>How it works</button>
          <button className="nav-link hide-mobile" onClick={()=>document.getElementById("pricing")?.scrollIntoView({behavior:"smooth"})}>Early Access</button>
          <button className="nav-link hide-mobile" onClick={()=>document.getElementById("compare")?.scrollIntoView({behavior:"smooth"})}>Compare</button>
          <button className="btn-cta" style={{padding:"10px 22px",fontSize:"0.875rem",marginLeft:8}} onClick={onLaunch}>Try Free →</button>
        </div>
      )}
      {appMode && <span className="badge badge-emerald">Live App</span>}
    </nav>
  );
}

function ScoreRing({score}) {
  const r=52,circ=2*Math.PI*r;
  const color=score>=70?"#0BC47F":score>=45?"#FF9F1C":"#FF4D6D";
  const offset=circ-(score/100)*circ;
  return (
    <div style={{width:130,height:130,display:"inline-flex",alignItems:"center",justifyContent:"center"}}>
      <svg width="130" height="130" viewBox="0 0 130 130">
        <circle cx="65" cy="65" r={r} fill="none" stroke="rgba(255,255,255,0.05)" strokeWidth="10"/>
        <circle cx="65" cy="65" r={r} fill="none" stroke={color} strokeWidth="10"
          strokeLinecap="round" strokeDasharray={circ} strokeDashoffset={offset}
          transform="rotate(-90 65 65)"
          style={{transition:"stroke-dashoffset 1.8s cubic-bezier(0.4,0,0.2,1)",filter:`drop-shadow(0 0 8px ${color}80)`}}/>
        <text x="65" y="62" textAnchor="middle" fill={color} style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"2rem"}}>{score}</text>
        <text x="65" y="80" textAnchor="middle" fill="var(--text3)" style={{fontFamily:"var(--fb)",fontSize:"0.65rem",fontWeight:500}}>/ 100</text>
      </svg>
    </div>
  );
}

function BarRow({item,delay=0}) {
  return (
    <div style={{marginBottom:14}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:7,alignItems:"center"}}>
        <span style={{fontSize:"0.875rem",fontWeight:500}}>{item.theme}</span>
        <span style={{fontSize:"0.78rem",fontWeight:700,color:item.color,fontFamily:"var(--fd)"}}>{item.pct}%</span>
      </div>
      <div className="prog-track">
        <div className="prog-fill" style={{width:`${item.pct}%`,background:`linear-gradient(90deg,${item.color}99,${item.color})`,boxShadow:`0 0 8px ${item.color}40`,animation:`bar-grow 1s ${delay}ms cubic-bezier(0.34,1.06,0.64,1) both`,"--w":`${item.pct}%`}}/>
      </div>
    </div>
  );
}

function UploadZone({onFile,file}) {
  const [drag,setDrag]=useState(false);
  const ref=useRef();
  const handle=f=>{if(f&&(f.name.endsWith(".csv")||f.name.endsWith(".xlsx")||f.name.endsWith(".xls")))onFile(f)};
  return (
    <div className={`upload-zone ${drag?"drag-over":"idle"}`}
      onClick={()=>ref.current.click()}
      onDragOver={e=>{e.preventDefault();setDrag(true)}}
      onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);handle(e.dataTransfer.files[0])}}>
      <input ref={ref} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>handle(e.target.files[0])}/>
      {file?(
        <div>
          <div style={{fontSize:"2.5rem",marginBottom:10}}>✅</div>
          <p style={{fontFamily:"var(--fd)",fontWeight:800,color:"var(--emerald)",fontSize:"1rem",marginBottom:4}}>{file.name}</p>
          <p style={{color:"var(--text3)",fontSize:"0.78rem"}}>{(file.size/1024).toFixed(1)} KB · Click to replace</p>
        </div>
      ):(
        <div>
          <div style={{fontSize:"3rem",marginBottom:14,filter:"drop-shadow(0 0 20px rgba(91,79,255,0.4))"}}>📂</div>
          <p style={{fontFamily:"var(--fd)",fontWeight:800,fontSize:"1.1rem",marginBottom:8}}>Drop your file here</p>
          <p style={{color:"var(--text2)",fontSize:"0.85rem",marginBottom:20,lineHeight:1.6}}>Supports <span style={{color:"var(--indigo2)"}}>CSV</span> and <span style={{color:"var(--indigo2)"}}>Excel</span> (.xlsx / .xls)</p>
          <span className="badge badge-indigo">Browse files →</span>
        </div>
      )}
    </div>
  );
}

function Report({data,onReset}) {
  const sb=data.sentiment==="Mostly Positive"?"badge-emerald":data.sentiment==="Mixed"?"badge-amber":"badge-coral";
  return (
    <div style={{position:"relative",zIndex:1}}>
      <div className="fu" style={{display:"flex",flexWrap:"wrap",gap:16,alignItems:"center",justifyContent:"space-between",marginBottom:36}}>
        <div>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:8}}>
            <h2 style={{fontSize:"1.75rem",fontWeight:900,letterSpacing:"-0.02em"}}>CustomerPulse Report</h2>
            <span className={`badge ${sb}`}>{data.sentiment}</span>
          </div>
          <p style={{color:"var(--text3)",fontSize:"0.82rem"}}>{data.reviewCount} reviews analyzed · {new Date().toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}</p>
        </div>
        <button className="btn-ghost" onClick={onReset}>← Back to Home</button>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:20}}>
        <div className="fu fu1" style={{display:"grid",gridTemplateColumns:"140px 1fr",gap:20}}>
          <div className="card card-glow" style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:10,padding:24}}>
            <ScoreRing score={data.score}/>
            <span className={`badge ${sb}`} style={{fontSize:"0.65rem"}}>{data.sentiment}</span>
          </div>
          <div className="card card-glow" style={{padding:28}}>
            <div className="sec-head">
              <div className="sec-icon" style={{background:"rgba(91,79,255,0.12)"}}>📋</div>
              <span className="sec-title">Executive Summary</span>
              <span className="sec-letter">A</span>
            </div>
            <p style={{color:"var(--text2)",lineHeight:1.8,fontSize:"0.92rem"}}>{data.summary}</p>
          </div>
        </div>
        <div className="report-grid fu fu2">
          <div className="card card-glow card-coral" style={{padding:24}}>
            <div className="sec-head">
              <div className="sec-icon" style={{background:"rgba(255,77,109,0.1)"}}>⚠️</div>
              <span className="sec-title">Top Complaints</span>
              <span className="sec-letter">B</span>
            </div>
            {data.complaints.map((c,i)=><BarRow key={c.theme} item={c} delay={i*80}/>)}
          </div>
          <div className="card card-glow card-emerald" style={{padding:24}}>
            <div className="sec-head">
              <div className="sec-icon" style={{background:"rgba(11,196,127,0.1)"}}>✨</div>
              <span className="sec-title">Positive Themes</span>
              <span className="sec-letter">C</span>
            </div>
            {data.positives.map((p,i)=><BarRow key={p.theme} item={p} delay={i*80}/>)}
          </div>
        </div>
        <div className="card card-glow card-indigo fu fu3" style={{padding:24}}>
          <div className="sec-head">
            <div className="sec-icon" style={{background:"rgba(91,79,255,0.1)"}}>💡</div>
            <span className="sec-title">Customer Requests</span>
            <span className="sec-letter">D</span>
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
            {data.requests.map(r=>(
              <span key={r} style={{padding:"10px 18px",borderRadius:"100px",background:"rgba(91,79,255,0.1)",border:"1px solid rgba(91,79,255,0.25)",color:"var(--indigo2)",fontSize:"0.875rem",fontWeight:600,fontFamily:"var(--fd)"}}>{r}</span>
            ))}
          </div>
        </div>
        <div className="card card-glow fu fu4" style={{padding:28}}>
          <div className="sec-head">
            <div className="sec-icon" style={{background:"rgba(255,159,28,0.1)"}}>🎯</div>
            <span className="sec-title">Recommended Actions</span>
            <span className="sec-letter">E</span>
          </div>
          <div style={{display:"flex",flexDirection:"column"}}>
            {data.actions.map((a,i)=>(
              <div key={i}>
                {i>0&&<div style={{height:1,background:"var(--border)",margin:"16px 0"}}/>}
                <div style={{display:"flex",gap:14,alignItems:"flex-start"}}>
                  <div style={{width:46,height:46,borderRadius:12,background:"rgba(255,159,28,0.08)",border:"1px solid rgba(255,159,28,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.2rem",flexShrink:0}}>{a.emoji}</div>
                  <div style={{flex:1}}>
                    <p style={{fontFamily:"var(--fd)",fontWeight:700,marginBottom:5,fontSize:"0.95rem"}}>{a.title}</p>
                    <p style={{color:"var(--text2)",fontSize:"0.86rem",lineHeight:1.7}}>{a.detail}</p>
                  </div>
                  <div style={{width:26,height:26,borderRadius:"50%",background:"rgba(255,159,28,0.1)",border:"1px solid rgba(255,159,28,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.75rem",color:"var(--amber)",fontWeight:800,fontFamily:"var(--fd)",flexShrink:0}}>{i+1}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
        <div className="fu fu5" style={{background:"linear-gradient(135deg,rgba(255,159,28,0.08) 0%,rgba(91,79,255,0.08) 100%)",border:"1px solid rgba(255,159,28,0.2)",borderRadius:"var(--r3)",padding:28,position:"relative",overflow:"hidden"}}>
          <div style={{position:"absolute",top:-30,right:-30,width:120,height:120,borderRadius:"50%",background:"rgba(255,159,28,0.08)",filter:"blur(30px)"}}/>
          <div className="sec-head">
            <div className="sec-icon" style={{background:"rgba(255,159,28,0.1)"}}>💰</div>
            <span className="sec-title" style={{color:"var(--amber2)"}}>Revenue Opportunity</span>
            <span className="sec-letter">F</span>
          </div>
          <p style={{lineHeight:1.8,fontSize:"0.93rem"}}>{data.opportunity}</p>
        </div>
        <div className="card" style={{padding:20,borderStyle:"dashed",textAlign:"center"}}>
          <p style={{color:"var(--text3)",fontSize:"0.75rem",fontWeight:700,fontFamily:"var(--fd)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:12}}>🔮 Coming Soon</p>
          <div style={{display:"flex",flexWrap:"wrap",gap:8,justifyContent:"center"}}>
            {["PDF Export","Save Reports","Google Reviews","Amazon Integration","Team Sharing","Yelp Import","Shopify Connect"].map(f=>(
              <span key={f} className="tag">{f}</span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function AppView({onBack}) {
  const [biz,setBiz]         = useState("");
  const [method,setMethod]   = useState("csv");
  const [file,setFile]       = useState(null);
  const [paste,setPaste]     = useState("");
  const [report,setReport]   = useState(null);
  const [error,setError]     = useState("");
  const [columns,setColumns] = useState([]);
  const [selCol,setSelCol]   = useState("");
  const [rows,setRows]       = useState([]);
  const [useAI,setUseAI]     = useState(false);
  const [loading,setLoading] = useState(false);
  const [loadMsg,setLoadMsg] = useState("Reading your reviews…");

  const RCOLS=["review","reviews","comment","comments","feedback","body","text","content","description","message"];
  const autoDetect=hs=>hs.find(h=>RCOLS.includes(h.toLowerCase().trim()))||null;

  const loadSample=()=>{
    const t=biz||"Restaurant";if(!biz)setBiz("Restaurant");
    setPaste((SAMPLES[t]||SAMPLES["Restaurant"]).join("\n"));setMethod("paste");
  };

  const handleFile=async f=>{
    setFile(f);setError("");setColumns([]);setSelCol("");setRows([]);
    try {
      if(f.name.endsWith(".csv")){
        Papa.parse(f,{header:true,skipEmptyLines:true,complete:res=>{
          setRows(res.data);
          const d=autoDetect(res.meta.fields||[]);
          d?setSelCol(d):setColumns(res.meta.fields||[]);
        },error:()=>setError("Could not parse CSV.")});
      } else {
        const buf=await f.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{defval:""});
        const headers=data.length?Object.keys(data[0]):[];
        setRows(data);
        const d=autoDetect(headers);d?setSelCol(d):setColumns(headers);
      }
    } catch{setError("Could not read file.");}
  };

  const analyze=async()=>{
    setError("");
    let reviews=[];
    if(method==="paste"){
      reviews=paste.split("\n").map(s=>s.trim()).filter(Boolean);
      if(reviews.length<2){setError("Please paste at least 2 reviews.");return;}
    } else {
      if(!file){setError("Please upload a file first.");return;}
      if(!rows.length){setError("File appears to be empty.");return;}
      if(!selCol){setError("Please select the review column.");return;}
      reviews=rows.map(r=>String(r[selCol]||"").trim()).filter(Boolean);
      if(reviews.length<2){setError("No review text found in selected column.");return;}
    }
    if(!biz){setError("Please select your business type.");return;}
    setLoading(true);
    const msgs=["Reading your reviews…","Identifying patterns…","Analyzing sentiment…","Generating your report…"];
    let i=0;const iv=setInterval(()=>{if(++i<msgs.length)setLoadMsg(msgs[i]);},900);
    try {
      const result=useAI?await aiAnalyze(reviews,biz):(await new Promise(r=>setTimeout(r,3200)),mockAnalyze(reviews,biz));
      setReport(result);
    } catch{setError("Analysis failed. "+(useAI?"Check connection or try mock mode.":"Please try again."));}
    finally{clearInterval(iv);setLoading(false);}
  };

  if(loading) return (
    <div style={{minHeight:"60vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:24}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3}}><div className="load-bar"/></div>
      <div className="spinner"/>
      <div style={{textAlign:"center"}}>
        <p style={{fontFamily:"var(--fd)",fontWeight:800,fontSize:"1.1rem",marginBottom:6}}>{loadMsg}</p>
        <p style={{color:"var(--text3)",fontSize:"0.85rem"}}>Usually takes under 5 seconds</p>
      </div>
    </div>
  );

  if(report) return <Report data={report} onReset={onBack}/>;

  return (
    <div style={{maxWidth:680,margin:"0 auto",padding:"44px 20px 60px"}}>
      <div className="fu" style={{display:"flex",alignItems:"center",gap:8,marginBottom:36}}>
        {["Business Type","Upload Reviews","Analyze"].map((s,i)=>(
          <div key={s} style={{display:"flex",alignItems:"center",gap:8}}>
            <span className={`step-dot${(i===0&&biz)||(i===1&&(file||paste))?" done":i===(biz?(file||paste)?2:1:0)?" active":""}`}>
              {((i===0&&biz)||(i===1&&(file||paste)))?"✓":i+1}
            </span>
            <span style={{fontSize:"0.78rem",color:"var(--text3)",fontWeight:500}}>{s}</span>
            {i<2&&<span style={{color:"var(--border2)",fontSize:"0.75rem",margin:"0 4px"}}>›</span>}
          </div>
        ))}
        <button className="btn-ghost" style={{marginLeft:"auto",fontSize:"0.75rem"}} onClick={()=>{const t=biz||"Restaurant";if(!biz)setBiz("Restaurant");setPaste((SAMPLES[t]||SAMPLES["Restaurant"]).join("\n"));setMethod("paste");}}>🎯 Sample data</button>
      </div>

      <h2 className="fu fu1" style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"1.8rem",letterSpacing:"-0.02em",marginBottom:32}}>Analyze your reviews</h2>

      <div className="card fu fu2" style={{marginBottom:16,padding:24}}>
        <p style={{fontSize:"0.72rem",fontWeight:800,letterSpacing:"0.1em",textTransform:"uppercase",color:"var(--text3)",marginBottom:14,fontFamily:"var(--fd)"}}>Step 1 — Business Type</p>
        <div style={{position:"relative"}}>
          <select value={biz} onChange={e=>setBiz(e.target.value)}>
            <option value="">Select your business type…</option>
            {["Restaurant","Retail Store","Ecommerce Product","Service Business","Other"].map(t=><option key={t} value={t}>{t}</option>)}
          </select>
          <span style={{position:"absolute",right:16,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:"var(--text3)"}}>▾</span>
        </div>
      </div>

      <div className="card fu fu3" style={{marginBottom:16,padding:24}}>
        <p style={{fontSize:"0.72rem",fontWeight:800,letterSpacing:"0.1em",textTransform:"uppercase",color:"var(--text3)",marginBottom:14,fontFamily:"var(--fd)"}}>Step 2 — Provide your reviews</p>
        <div style={{display:"flex",gap:10,marginBottom:20}}>
          {[{k:"csv",e:"📄",l:"CSV File"},{k:"excel",e:"📊",l:"Excel File"},{k:"paste",e:"📋",l:"Paste Text"}].map(m=>(
            <button key={m.k} className={`im-tab${method===m.k?" active":""}`} onClick={()=>setMethod(m.k)}>
              <span className="im-emoji">{m.e}</span>
              <span className="im-label">{m.l}</span>
            </button>
          ))}
        </div>
        {method==="paste"?(
          <div>
            <p style={{fontSize:"0.78rem",color:"var(--text3)",marginBottom:10}}>One review per line. Minimum 2 reviews.</p>
            <textarea value={paste} onChange={e=>setPaste(e.target.value)} placeholder={"The food was amazing!\nService was too slow.\nBest place in town."}/>
            {paste&&<p style={{fontSize:"0.72rem",color:"var(--text3)",marginTop:8}}>{paste.split("\n").filter(Boolean).length} reviews detected</p>}
          </div>
        ):(
          <div>
            <UploadZone onFile={f=>{setFile(f);handleFile(f);}} file={file}/>
            {columns.length>0&&!selCol&&(
              <div style={{marginTop:14,padding:"16px 18px",background:"rgba(255,159,28,0.05)",border:"1px solid rgba(255,159,28,0.2)",borderRadius:"var(--r2)"}}>
                <p style={{fontSize:"0.82rem",color:"var(--amber2)",marginBottom:12,fontWeight:500}}>⚠️ Select which column contains your reviews:</p>
                <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                  {columns.map(c=><button key={c} className={`col-sel-btn${selCol===c?" active":""}`} onClick={()=>setSelCol(c)}>{c}</button>)}
                </div>
              </div>
            )}
            {selCol&&<p style={{fontSize:"0.73rem",color:"var(--emerald)",marginTop:10,fontWeight:600}}>✓ Column: <strong>"{selCol}"</strong>{rows.length>0&&` · ${rows.length} rows`}</p>}
          </div>
        )}
      </div>

      <div className="card fu fu4" style={{marginBottom:16,padding:"14px 20px",display:"flex",alignItems:"center",gap:14}}>
        <button className="toggle-track" style={{background:useAI?"var(--indigo)":"rgba(255,255,255,0.08)"}} onClick={()=>setUseAI(!useAI)}>
          <div className="toggle-thumb" style={{left:useAI?"23px":"3px"}}/>
        </button>
        <div>
          <p style={{fontFamily:"var(--fd)",fontWeight:700,fontSize:"0.875rem",marginBottom:2}}>{useAI?"🤖 Real AI Analysis (Claude)":"⚡ Smart Mock Analysis"}</p>
          <p style={{fontSize:"0.75rem",color:"var(--text3)"}}>{useAI?"Uses Anthropic API":"Instant results, no API key needed"}</p>
        </div>
        <span className="badge" style={{marginLeft:"auto",...(useAI?{background:"var(--indigo3)",color:"var(--indigo2)",border:"1px solid rgba(91,79,255,0.25)"}:{background:"var(--amber3)",color:"var(--amber2)",border:"1px solid rgba(255,159,28,0.25)"})}}>{useAI?"AI":"MOCK"}</span>
      </div>

      {error&&<div className="error-box fu" style={{marginBottom:16}}>⚠️ <span>{error}</span></div>}

      <button className="btn-cta fu fu5"
        style={{width:"100%",justifyContent:"center",padding:"18px",fontSize:"1.05rem",borderRadius:"var(--r2)"}}
        onClick={analyze}
        disabled={!biz||(method!=="paste"&&!file)||(method==="paste"&&!paste.trim())}>
        Analyze Feedback →
      </button>
      <p style={{textAlign:"center",fontSize:"0.75rem",color:"var(--text3)",marginTop:14}}>Your data is never stored or shared.</p>
    </div>
  );
}

function WaitlistCard() {
  const [email,setEmail]   = useState("");
  const [status,setStatus] = useState("idle");

  const submit = () => {
    if(!email||!email.includes("@")){setStatus("error");return;}
    setStatus("success");
    setEmail("");
  };

  return (
    <div className="card card-featured" style={{padding:36,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-40,right:-40,width:160,height:160,borderRadius:"50%",background:"rgba(255,159,28,0.06)",filter:"blur(40px)"}}/>
      <div style={{position:"relative"}}>
        <span className="badge badge-amber" style={{marginBottom:20,display:"inline-flex"}}>🔒 Premium Coming Soon</span>
        <h3 style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"1.4rem",letterSpacing:"-0.02em",marginBottom:12,lineHeight:1.2}}>
          Be first in line for<br/>premium features.
        </h3>
        <p style={{color:"var(--text2)",fontSize:"0.9rem",lineHeight:1.7,marginBottom:28}}>
          We're building PDF exports, saved reports, Google Reviews import, monthly trend tracking, and team collaboration. Join the list and get early access + a special launch discount.
        </p>
        <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:20}}>
          {["PDF report export","Save & compare reports over time","Google & Amazon review import","Monthly trend alerts","Team sharing & collaboration"].map(f=>(
            <div key={f} style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:18,height:18,borderRadius:"50%",background:"rgba(255,159,28,0.15)",border:"1px solid rgba(255,159,28,0.3)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.6rem",color:"var(--amber)",flexShrink:0}}>✓</div>
              <span style={{fontSize:"0.82rem",color:"var(--text2)"}}>{f}</span>
            </div>
          ))}
        </div>
        <div style={{height:1,background:"var(--border)",margin:"24px 0"}}/>
        {status==="success"?(
          <div style={{padding:"16px 20px",background:"rgba(11,196,127,0.08)",border:"1px solid rgba(11,196,127,0.25)",borderRadius:"var(--r2)",textAlign:"center"}}>
            <p style={{fontFamily:"var(--fd)",fontWeight:800,color:"var(--emerald)",fontSize:"1rem",marginBottom:4}}>🎉 You're on the list!</p>
            <p style={{fontSize:"0.82rem",color:"var(--text2)"}}>We'll email you the moment premium launches — with an exclusive early-bird offer.</p>
          </div>
        ):(
          <div>
            <p style={{fontSize:"0.78rem",fontWeight:700,fontFamily:"var(--fd)",letterSpacing:"0.06em",textTransform:"uppercase",color:"var(--text3)",marginBottom:10}}>Get early access</p>
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              <input
                type="email"
                value={email}
                onChange={e=>{setEmail(e.target.value);setStatus("idle")}}
                placeholder="your@email.com"
                style={{flex:1,minWidth:0,padding:"13px 16px",borderRadius:"var(--r2)",background:"rgba(255,255,255,0.05)",border:`1px solid ${status==="error"?"var(--coral)":"var(--border2)"}`,color:"var(--text)",fontFamily:"var(--fb)",fontSize:"0.9rem",outline:"none"}}
              />
              <button className="btn-plan-cta" style={{width:"auto",padding:"13px 22px",borderRadius:"var(--r2)",whiteSpace:"nowrap"}} onClick={submit}>
                Notify Me →
              </button>
            </div>
            {status==="error"&&<p style={{fontSize:"0.75rem",color:"var(--coral)",marginTop:8}}>Please enter a valid email address.</p>}
            <p style={{fontSize:"0.72rem",color:"var(--text3)",marginTop:10}}>No spam. Unsubscribe any time. We'll only email when it's ready.</p>
          </div>
        )}
      </div>
    </div>
  );
}

export default function CustomerPulse() {
  const [view,setView] = useState("home");

  if(view==="app") return (
    <div style={{minHeight:"100vh",position:"relative"}}>
      <style>{CSS}</style>
      <Orbs/>
      <Nav onLogoClick={()=>setView("home")} onLaunch={()=>setView("app")} appMode={true}/>
      <div style={{position:"relative",zIndex:1}}>
        <AppView onBack={()=>setView("home")}/>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh"}}>
      <style>{CSS}</style>
      <Orbs/>
      <Nav onLogoClick={()=>setView("home")} onLaunch={()=>setView("app")}/>

      {/* ── HERO ── */}
      <section style={{padding:"120px 24px 100px",textAlign:"center",position:"relative",zIndex:1}}>
        <div style={{maxWidth:820,margin:"0 auto"}}>
          <div className="badge badge-indigo fu" style={{marginBottom:28}}>✦ AI-Powered Review Intelligence</div>
          <h1 className="fu fu1" style={{fontSize:"clamp(2.8rem,7vw,5rem)",fontWeight:900,lineHeight:1.05,letterSpacing:"-0.03em",marginBottom:28}}>
            Your customers are<br/>telling you how to<br/>
            <span className="shimmer-text">grow your business.</span>
          </h1>
          <p className="fu fu2" style={{fontSize:"1.15rem",color:"var(--text2)",lineHeight:1.8,marginBottom:14,maxWidth:560,margin:"0 auto 14px"}}>
            Every review is a data point. CustomerPulse turns hundreds of reviews into a single clear report — with the exact actions that will move the needle.
          </p>
          <p className="fu fu2" style={{fontSize:"0.95rem",color:"var(--text3)",lineHeight:1.7,marginBottom:52,maxWidth:480,margin:"0 auto 52px"}}>
            Built for restaurant owners, retailers, ecommerce sellers, and service businesses who are tired of guessing.
          </p>
          <div className="fu fu3" style={{display:"flex",gap:14,flexWrap:"wrap",justifyContent:"center",marginBottom:64}}>
            <button className="btn-cta" onClick={()=>setView("app")}>Start Free — No Credit Card →</button>
            <button className="btn-secondary" onClick={()=>setView("app")}>See a sample report</button>
          </div>
          <div className="fu fu4" style={{display:"flex",gap:28,flexWrap:"wrap",justifyContent:"center"}}>
            {[["500+","Business owners"],["2 min","Average time to report"],["$0","To get started"],["6","Report sections"]].map(([n,l])=>(
              <div key={l} style={{textAlign:"center"}}>
                <p style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"1.6rem",color:"var(--text)",letterSpacing:"-0.02em"}}>{n}</p>
                <p style={{fontSize:"0.78rem",color:"var(--text3)",marginTop:2}}>{l}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      <div className="divider"/>

      {/* ── PROBLEM ── */}
      <section id="problem" style={{padding:"100px 24px",position:"relative",zIndex:1}}>
        <div className="section-inner">
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:60,alignItems:"center"}}>
            <div>
              <p className="section-label">The Problem</p>
              <h2 className="section-title">You're drowning in reviews.<br/>Starving for <span className="gradient-text">insight.</span></h2>
              <p style={{color:"var(--text2)",lineHeight:1.8,fontSize:"1rem",marginBottom:32}}>
                Every platform pushes customers to leave feedback. Google, Yelp, Amazon, TripAdvisor. Business owners read them one by one, feel good or bad, and move on.
              </p>
              <p style={{color:"var(--text2)",lineHeight:1.8,fontSize:"1rem"}}>
                There's no tool simple enough for a restaurant owner or small retailer to upload a file and get a clear action plan in 60 seconds. Until now.
              </p>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {[
                {icon:"😤",title:"Reading reviews one by one",sub:"Hours wasted with no clear takeaway"},
                {icon:"📊",title:"Exporting to spreadsheets",sub:"Complex formulas that still don't give answers"},
                {icon:"💸",title:"Hiring consultants",sub:"$5,000+ reports that take weeks to arrive"},
                {icon:"🤷",title:"Making decisions on gut feel",sub:"Fixing the wrong things, missing what matters"},
              ].map(({icon,title,sub})=>(
                <div key={title} style={{display:"flex",gap:14,alignItems:"center",padding:"16px 20px",background:"rgba(255,77,109,0.04)",border:"1px solid rgba(255,77,109,0.12)",borderRadius:"var(--r2)"}}>
                  <span style={{fontSize:"1.4rem",width:40,textAlign:"center"}}>{icon}</span>
                  <div>
                    <p style={{fontFamily:"var(--fd)",fontWeight:700,fontSize:"0.9rem",marginBottom:3}}>{title}</p>
                    <p style={{fontSize:"0.8rem",color:"var(--text2)"}}>{sub}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </section>

      <div className="divider"/>

      {/* ── HOW IT WORKS ── */}
      <section id="how" style={{padding:"100px 24px",position:"relative",zIndex:1}}>
        <div className="section-inner" style={{textAlign:"center"}}>
          <p className="section-label">How It Works</p>
          <h2 className="section-title" style={{margin:"0 auto 16px"}}>Three steps to a full<br/><span className="gradient-text">business intelligence report.</span></h2>
          <p style={{color:"var(--text2)",lineHeight:1.75,marginBottom:64,maxWidth:480,margin:"0 auto 64px"}}>No training. No setup. No data team required.</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:20,textAlign:"left"}}>
            {[
              {n:"01",icon:"📂",title:"Upload your reviews",body:"Drop a CSV or Excel export from Google, Yelp, Amazon, or any review platform. Or paste reviews directly. We handle the parsing automatically.",color:"var(--indigo)"},
              {n:"02",icon:"🧠",title:"AI analyzes the patterns",body:"Our AI reads every review, identifies recurring themes, measures sentiment, and separates signal from noise — in under 60 seconds.",color:"var(--amber)"},
              {n:"03",icon:"🎯",title:"Get your action report",body:"A clear 6-section report: sentiment score, top complaints, positive themes, customer requests, 5 specific actions, and your biggest revenue opportunity.",color:"var(--emerald)"},
            ].map(({n,icon,title,body,color})=>(
              <div key={n} className="card card-glow" style={{padding:28}}>
                <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:20}}>
                  <span style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"0.75rem",color,opacity:0.6}}>{n}</span>
                  <div style={{width:44,height:44,borderRadius:12,background:`${color}15`,border:`1px solid ${color}30`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.3rem"}}>{icon}</div>
                </div>
                <h3 style={{fontFamily:"var(--fd)",fontWeight:800,fontSize:"1.1rem",marginBottom:10}}>{title}</h3>
                <p style={{color:"var(--text2)",fontSize:"0.875rem",lineHeight:1.7}}>{body}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      <div className="divider"/>

      {/* ── COMPARE ── */}
      <section id="compare" style={{padding:"100px 24px",position:"relative",zIndex:1}}>
        <div className="section-inner">
          <div style={{textAlign:"center",marginBottom:56}}>
            <p className="section-label">Why CustomerPulse</p>
            <h2 className="section-title" style={{margin:"0 auto 16px"}}>The only tool built<br/><span className="gradient-text">for business owners, not analysts.</span></h2>
            <p style={{color:"var(--text2)",lineHeight:1.75,maxWidth:500,margin:"0 auto"}}>Every other option is either too complex, too expensive, or built for enterprise teams. We built CustomerPulse for the person actually running the business.</p>
          </div>
          <div className="card" style={{padding:0,overflow:"hidden"}}>
            <table className="comp-table">
              <thead>
                <tr>
                  <th style={{textAlign:"left",width:"30%"}}></th>
                  <th className="comp-head-us">CustomerPulse</th>
                  <th>Generic AI Tools</th>
                  <th className="hide-mobile">Spreadsheets</th>
                  <th className="hide-mobile">Consultants</th>
                </tr>
              </thead>
              <tbody>
                {[
                  ["Upload & get instant report","✓","✗","✗","✗"],
                  ["Business-specific insights","✓","✗","✗","✓"],
                  ["Actionable next steps","✓","~","✗","✓"],
                  ["Revenue opportunity analysis","✓","✗","✗","✓"],
                  ["Under 2 minutes","✓","✓","✗","✗"],
                  ["Free to start","✓","~","✓","✗"],
                  ["No technical skills needed","✓","~","✗","✓"],
                  ["CSV & Excel upload","✓","✗","✓","✗"],
                ].map(([feature,...vals])=>(
                  <tr key={feature}>
                    <td style={{fontWeight:500}}>{feature}</td>
                    {vals.map((v,i)=>(
                      <td key={i} className={i===0?"comp-col-us":""}>
                        <span style={{
                          color:v==="✓"?i===0?"var(--indigo2)":"var(--emerald)":v==="~"?"var(--amber2)":"var(--text3)",
                          fontWeight:v==="✓"?700:400,fontSize:"1rem"
                        }}>{v}</span>
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </section>

      <div className="divider"/>

      {/* ── FREE TRIAL + EMAIL CAPTURE ── */}
      <section id="pricing" style={{padding:"100px 24px",position:"relative",zIndex:1}}>
        <div className="section-inner">
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:48,alignItems:"center",maxWidth:1000,margin:"0 auto"}}>

            <div>
              <p className="section-label">Free Access</p>
              <h2 className="section-title">Try it free.<br/><span className="gradient-text">No strings attached.</span></h2>
              <p style={{color:"var(--text2)",lineHeight:1.8,fontSize:"1rem",marginBottom:32}}>
                Upload your reviews right now and get your first business intelligence report in under 2 minutes. No credit card. No account needed. Just results.
              </p>
              <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:36}}>
                {[
                  {icon:"📊",text:"Instant sentiment score from 1–100"},
                  {icon:"⚠️",text:"Your top complaints ranked by frequency"},
                  {icon:"✨",text:"What customers actually love about you"},
                  {icon:"🎯",text:"3 specific actions you can take this week"},
                  {icon:"💰",text:"Your single biggest revenue opportunity"},
                ].map(({icon,text})=>(
                  <div key={text} style={{display:"flex",alignItems:"center",gap:12}}>
                    <div style={{width:34,height:34,borderRadius:9,background:"var(--glass2)",border:"1px solid var(--border2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.95rem",flexShrink:0}}>{icon}</div>
                    <span style={{fontSize:"0.9rem",color:"var(--text2)"}}>{text}</span>
                  </div>
                ))}
              </div>
              <button className="btn-cta" onClick={()=>setView("app")}>Analyze My Reviews Free →</button>
              <p style={{fontSize:"0.75rem",color:"var(--text3)",marginTop:12}}>No account needed · Results in under 2 minutes</p>
            </div>

            <WaitlistCard/>

          </div>
        </div>
      </section>

      <div className="divider"/>

      {/* ── VISION ── */}
      <section style={{padding:"100px 24px",position:"relative",zIndex:1}}>
        <div className="section-inner">
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:60,alignItems:"center"}}>
            <div>
              <p className="section-label">Our Mission</p>
              <h2 className="section-title">Every business owner deserves a<br/><span className="gradient-text">data team in their pocket.</span></h2>
              <p style={{color:"var(--text2)",lineHeight:1.8,fontSize:"1rem",marginBottom:20}}>
                The big chains have analytics departments. Enterprise brands hire consultants. But the restaurant owner with 3 locations, the Amazon seller doing $30k a month, the local retailer — they've been left out.
              </p>
              <p style={{color:"var(--text2)",lineHeight:1.8,fontSize:"1rem",marginBottom:32}}>
                CustomerPulse exists to close that gap. AI makes it possible to turn 500 reviews into a one-page action plan that any business owner can read and act on today. That's what we're building.
              </p>
              <button className="btn-cta" onClick={()=>setView("app")}>Try it free →</button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              {[
                {icon:"🔮",title:"Coming next",items:["Google Business Profile import","Amazon review sync","Monthly trend reports","Competitor benchmarking"]},
                {icon:"🚀",title:"The roadmap",items:["Shopify & Etsy integration","Yelp & TripAdvisor import","Multi-language support","Team alerts & notifications"]},
              ].map(({icon,title,items})=>(
                <div key={title} className="card" style={{padding:24}}>
                  <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
                    <span style={{fontSize:"1.2rem"}}>{icon}</span>
                    <p style={{fontFamily:"var(--fd)",fontWeight:700}}>{title}</p>
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                    {items.map(item=><span key={item} className="tag">{item}</span>)}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </section>

      <div className="divider"/>

      {/* ── CTA FOOTER ── */}
      <section style={{padding:"100px 24px",textAlign:"center",position:"relative",zIndex:1}}>
        <div style={{maxWidth:640,margin:"0 auto"}}>
          <h2 style={{fontSize:"clamp(2rem,5vw,3.2rem)",fontWeight:900,letterSpacing:"-0.03em",marginBottom:20,lineHeight:1.15}}>
            Stop guessing.<br/><span className="shimmer-text">Start knowing.</span>
          </h2>
          <p style={{color:"var(--text2)",lineHeight:1.8,marginBottom:40,fontSize:"1.05rem"}}>
            Upload your reviews right now. Your first report is free and takes under 2 minutes.
          </p>
          <button className="btn-cta" style={{fontSize:"1.1rem",padding:"18px 44px"}} onClick={()=>setView("app")}>
            Analyze My Reviews Free →
          </button>
          <p style={{color:"var(--text2)",fontSize:"0.78rem",marginTop:16}}>No credit card · No setup · Results in under 2 minutes</p>
        </div>
      </section>

      <footer style={{borderTop:"1px solid var(--border)",padding:"32px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:16,position:"relative",zIndex:1}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:28,height:28,borderRadius:7,background:"linear-gradient(135deg,var(--indigo),var(--indigo2))",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.85rem"}}>📊</div>
          <span style={{fontFamily:"var(--fd)",fontWeight:900,fontSize:"0.95rem"}}>CustomerPulse</span>
        </div>
        <p style={{color:"var(--text2)",fontSize:"0.78rem"}}>© 2026 CustomerPulse · Built for business owners who mean business.</p>
        <div style={{display:"flex",gap:20}}>
          {["Privacy","Terms","Contact"].map(l=>(
            <button key={l} style={{background:"none",border:"none",color:"var(--text2)",fontSize:"0.78rem",cursor:"pointer"}}>{l}</button>
          ))}
        </div>
      </footer>
    </div>
  );
}
