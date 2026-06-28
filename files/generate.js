'use strict';
// generate.js — Astonish Commercial Real Estate
// Exports: generateDocx(data) → Promise<Buffer>

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, LevelFormat,
} = require('docx');

const C = { PINK:'E8007D', PINKLT:'FF3DA0', CYAN:'00AEEF', NAVY:'0A0F1E', GRAY:'F6F8FC', BORDER:'D0D7E6', WHITE:'FFFFFF' };
const LABELS = { lease:'Exclusive Right to Lease', sell:'Exclusive Right to Sell', sell_lease:'Exclusive Right to Sell and Lease', auction:'Exclusive Right to Sell (Auction)' };
const EXP_CATS = ['Real Estate Taxes & Assessments','Common Area Maintenance','RE Tax & Insurance Increase','Insurance','Janitorial','Water','Sewer','Gas','Electrical','Garbage Removal','Heating','Air Conditioning','Interior Maintenance','Exterior Maintenance','Structure','Parking Lot','Landscape','Other'];

async function generateDocx(data, sections = {}) { return Packer.toBuffer(buildDoc(data, sections)); }
module.exports = { generateDocx };

// ─── Primitive helpers ────────────────────────────────────────────────────────
const tx = (text, opts={}) => new TextRun({ text, font:'Arial', size:18, ...opts });
const blk = () => new Paragraph({ spacing:{before:40,after:40}, children:[tx('')] });
const pct = v => v ? `${v}%` : '______%';
const fmtDate = str => { if (!str) return ''; return new Date(str+'T00:00:00').toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'}); };
const joinAddr = (a,b,c,d) => { const l2=[b,c,d].filter(Boolean).join(' '); return [a,l2].filter(Boolean).join(', '); };
const allBdrs = (color=C.BORDER) => { const b={style:BorderStyle.SINGLE,size:1,color}; return {top:b,bottom:b,left:b,right:b}; };
const noBdrs = () => { const b={style:BorderStyle.NONE,size:0,color:C.WHITE}; return {top:b,bottom:b,left:b,right:b}; };
const cellPad = () => ({top:80,bottom:80,left:120,right:120});

function p(text, opts={}) {
  const {bold=false,before=60,after=60,size=18,color=undefined,italics=false}=opts;
  return new Paragraph({ spacing:{before,after}, children:[tx(text,{bold,size,color,italics})] });
}

function sigLine(label) {
  return [
    new Paragraph({ border:{bottom:{style:BorderStyle.SINGLE,size:6,color:'333333'}}, spacing:{before:200,after:28}, children:[tx('')] }),
    new Paragraph({ spacing:{before:0,after:120}, children:[tx(label,{size:15,color:'888888'})] }),
  ];
}

function banner(text, bg, fg=C.WHITE) {
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[9360], rows:[new TableRow({ children:[new TableCell({
    borders:allBdrs(bg), shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:160,right:160}, width:{size:9360,type:WidthType.DXA},
    children:[new Paragraph({ children:[tx(text.toUpperCase(),{bold:true,color:fg,size:19})] })],
  })]})]});
}

const sectionHdr = text => banner(text, C.CYAN);
const legalHdr   = text => banner(text, C.NAVY, C.PINKLT);

function dataRow(lbl, val, w1=2600, w2=6760) {
  return new TableRow({ children:[
    new TableCell({ borders:allBdrs(), shading:{fill:C.GRAY,type:ShadingType.CLEAR}, margins:{top:70,bottom:70,left:120,right:120}, width:{size:w1,type:WidthType.DXA}, children:[new Paragraph({ children:[tx(lbl,{bold:true,size:17,color:'444444'})] })] }),
    new TableCell({ borders:allBdrs(), margins:{top:70,bottom:70,left:120,right:120}, width:{size:w2,type:WidthType.DXA}, children:[new Paragraph({ children:[tx(val||'—',{size:17})] })] }),
  ]});
}

// ─── Expense table (uses filled data from form) ───────────────────────────────
function expenseTable(expenses) {
  const cols=[3100,1570,1570,1570,1550];
  const expMap={};
  (expenses||[]).forEach(e=>{ if(e&&e.cat) expMap[e.cat]=e; });

  const hCell=(text,w)=>new TableCell({ borders:allBdrs(), shading:{fill:'E8EEF8',type:ShadingType.CLEAR}, margins:{top:60,bottom:60,left:80,right:80}, width:{size:w,type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({alignment:AlignmentType.CENTER,children:[tx(text,{bold:true,size:15,color:'333333'})]})] });
  const dCell=(text,w,shade)=>new TableCell({ borders:allBdrs(), shading:{fill:shade?'F7F9FF':C.WHITE,type:ShadingType.CLEAR}, margins:{top:55,bottom:55,left:80,right:80}, width:{size:w,type:WidthType.DXA},
    children:[new Paragraph({alignment:text==='✓'||text==='☐'?AlignmentType.CENTER:AlignmentType.LEFT,children:[tx(text,{size:16,bold:text==='✓',color:text==='✓'?'16A34A':'333333'})]})] });

  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:cols, rows:[
    new TableRow({ tableHeader:true, children:[hCell('Expense Category',cols[0]),hCell('Landlord Pays',cols[1]),hCell('Tenant Pays',cols[2]),hCell('Pro-Rata',cols[3]),hCell('Prior Yr PSF',cols[4])] }),
    ...EXP_CATS.map((cat,i)=>{ const row=expMap[cat]||{}; const s=i%2===0;
      return new TableRow({ children:[dCell(cat,cols[0],s),dCell(row.landlord?'✓':'☐',cols[1],s),dCell(row.tenant?'✓':'☐',cols[2],s),dCell(row.prorata?'✓':'☐',cols[3],s),dCell(row.psf?`$${row.psf}`:'',cols[4],s)] }); }),
  ]});
}

// ─── Document builder ─────────────────────────────────────────────────────────
function buildDoc(d, sections = {}) {
  const t=d.type, isLease=['lease','sell_lease'].includes(t), isSell=['sell','sell_lease','auction'].includes(t);
  const s = sections; // shorthand for sections config
  const today=new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'});
  const ownerLbl=t==='lease'?'Lessor / Owner':'Owner';
  const agentAddr=joinAddr(d.agentAddress,d.agentCity,d.agentState,d.agentZip);
  const ownerAddr=joinAddr(d.ownerAddress,d.ownerCity,d.ownerState,d.ownerZip);
  const propFull =joinAddr(d.propAddr,d.propCity,d.propState,d.propZip);
  const signs=[d.signBldg&&'Building Sign(s)',d.signAsph&&'Asphalt/Rebar Spike Sign(s)',d.signFnce&&'Fence Sign(s)',d.signWndw&&'Window Sign(s)',d.signYard&&'Yard Sign(s)'].filter(Boolean);
  const fm=d.fillMode==='owner', lp=fm?'______%':pct(d.leaseComm), rp=fm?'______%':pct(d.renewComm), sp=pct(d.saleComm);
  const bas=fm?'[ ] Gross  [ ] Net':(d.commBasis==='net'?'net':'gross');
  const typeName={lease:'Lease',sell:'Sell',sell_lease:'Sell and Lease',auction:'Sell (Auction)'}[t];

  const ch=[];
  ch.push(p(`This Exclusive Right to ${typeName} ("Agreement"), dated as of the date of last signature ("Effective Date"), is entered into by and between Owner and Astonish, LLC, an Ohio limited liability company (together with its affiliates, successors and assigns, "Brokerage Firm").`,{before:0,after:180}));

  // PARTIES
  ch.push(sectionHdr('Parties')); ch.push(blk());
  const half=4680;
  ch.push(new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[half,half], rows:[
    new TableRow({ children:[
      new TableCell({ borders:allBdrs(C.CYAN), shading:{fill:C.CYAN,type:ShadingType.CLEAR}, margins:{top:60,bottom:60,left:120,right:120}, width:{size:half,type:WidthType.DXA}, children:[new Paragraph({children:[tx(ownerLbl.toUpperCase(),{bold:true,color:C.WHITE,size:17})]})] }),
      new TableCell({ borders:allBdrs(C.CYAN), shading:{fill:C.CYAN,type:ShadingType.CLEAR}, margins:{top:60,bottom:60,left:120,right:120}, width:{size:half,type:WidthType.DXA}, children:[new Paragraph({children:[tx('BROKERAGE FIRM',{bold:true,color:C.WHITE,size:17})]})] }),
    ]}),
    new TableRow({ children:[
      new TableCell({ borders:allBdrs(), margins:cellPad(), width:{size:half,type:WidthType.DXA}, children:[
        new Paragraph({children:[tx(d.ownerName||'___________________________',{bold:true,size:18})]}),
        ...(d.ownerContact?[new Paragraph({children:[tx(`Contact: ${d.ownerContact}`,{size:17})]})]:[]),
        ...(ownerAddr?[new Paragraph({children:[tx(ownerAddr,{size:17})]})]:[]),
        ...(d.ownerPhone?[new Paragraph({children:[tx(`Phone: ${d.ownerPhone}`,{size:17})]})]:[]),
        ...(d.ownerCell?[new Paragraph({children:[tx(`Cell: ${d.ownerCell}`,{size:17})]})]:[]),
        ...(d.ownerEmail?[new Paragraph({children:[tx(`Email: ${d.ownerEmail}`,{size:17})]})]:[]),
      ]}),
      new TableCell({ borders:allBdrs(), margins:cellPad(), width:{size:half,type:WidthType.DXA}, children:[
        new Paragraph({children:[tx('Astonish LLC',{bold:true,size:18})]}),
        new Paragraph({children:[tx(`Agent: ${d.agentName||(s.defaults?.agent_name||'Michael Bergman')}`,{size:17})]}),
        new Paragraph({children:[tx(`Title: ${d.agentTitle||(s.defaults?.agent_title||'Broker')}`,{size:17})]}),
        new Paragraph({children:[tx(agentAddr||(s.defaults?.agent_address||'9918 Carver Rd., Suite 101, Cincinnati OH 45242'),{size:17})]}),
        new Paragraph({children:[tx(`Phone: ${d.agentPhone||(s.defaults?.agent_phone||'513.334.3624')}`,{size:17})]}),
        ...(d.agentCell?[new Paragraph({children:[tx(`Cell: ${d.agentCell}`,{size:17})]})]:[]),
        new Paragraph({children:[tx(`Email: ${d.agentEmail||(s.defaults?.agent_email||'info@astonishcommercial.com')}`,{size:17})]}),
      ]}),
    ]}),
  ]}));
  ch.push(blk());

  // PROPERTY
  ch.push(sectionHdr('Property'));
  const pRows=[dataRow('Property Address',propFull||'—')];
  if(d.propApn) pRows.push(dataRow('APN / Parcel Number(s)',d.propApn));
  if(t==='auction'&&d.prop2) pRows.push(dataRow('Property 2',d.prop2));
  if(t==='auction'&&d.prop3) pRows.push(dataRow('Property 3',d.prop3));
  const propertyIncludesText = t==='lease' ? (s.type_specific?.lease?.property_includes || 'Broker is authorized to market the Property including all buildings, improvements, fixtures, easements, rights-of-way, leases, rents, security deposits, licenses, permits, transferable warranties, and personal property used in operations.') : (s.type_specific?.[t]?.property_includes || "Except as specifically provided in this Agreement, Broker is to market the Property including: (a) all buildings, improvements, and fixtures; (b) all rights, privileges, and appurtenances pertaining to the Property, including Owner's right, title, and interest in any easements and rights-of-way; (c) Owner's interest in all leases, rents, and security deposits; (d) Owner's interest in all licenses, permits and transferable warranties related to the Property; and (e) all Owner's personal property located on the Property that is used in connection with the Property's operations except as set out on the attached Listing Agreement Addendum.");
  pRows.push(dataRow('Property Includes', propertyIncludesText));
  ch.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2600,6760],rows:pRows}));
  ch.push(blk());

  // TERM
  ch.push(sectionHdr('Listing Term'));
  const tRows=[dataRow('Start Date',fmtDate(d.tStart)||'—'),dataRow('End Date',fmtDate(d.tEnd)||'—')];
  if(t==='auction') tRows.push(dataRow('Auction Agreement Execution Date',fmtDate(d.auctionDate)||'—'));
  ch.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2600,6760],rows:tRows}));
  ch.push(blk());

  // SALE TERMS
  if(isSell){
    ch.push(sectionHdr('Sale Terms'));
    const sRows=[];
    if(t!=='auction') sRows.push(dataRow('Listing / Sale Price',d.listPrice||'—'));
    sRows.push(dataRow('Sale Commission',`${sp} of gross sales price${t==='auction'?" (excl. buyer's premium)":''}` ));
    if(t==='auction'){
      sRows.push(dataRow('Reserve Price – Property 1',d.res1||'Hidden Reserve Price'));
      if(d.prop2) sRows.push(dataRow('Reserve Price – Property 2',d.res2||'Hidden Reserve Price'));
      if(d.prop3) sRows.push(dataRow('Reserve Price – Property 3',d.res3||'Hidden Reserve Price'));
      sRows.push(dataRow('Termination Fee',`Greater of ${d.termFee||'3'}% of Reserve Price or $20,000.00`));
    }
    ch.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2600,6760],rows:sRows}));
    ch.push(blk());
  }

  // LEASE TERMS + EXPENSE TABLE
  if(isLease){
    const ownerSfx=fm?' [Owner to complete when signing]':'';
    ch.push(sectionHdr('Lease Terms'));
    ch.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2600,6760],rows:[
      dataRow('Lease Type',     fm?'____________________'+ownerSfx:(d.leaseType||'—')),
      dataRow('Desired Term',   fm?'____________________'+ownerSfx:(d.leaseTerm||'—')),
      dataRow('Base Lease Rate',fm?'____________________'+ownerSfx:(d.baseRate||'—')),
      dataRow('NNN Expenses',   fm?'____________________'+ownerSfx:(d.nnn||'—')),
    ]}));
    ch.push(blk());
    ch.push(p('Expense Allocation',{bold:true,before:60,after:40}));
    ch.push(p('Check who pays each expense. Fields left blank may be completed at signing by either party.',{before:0,after:60,size:16,color:'555555'}));
    ch.push(expenseTable(d.expenses||[]));
    ch.push(blk());
  }

  // COMMISSION
  ch.push(sectionHdr('Commission')); ch.push(blk());
  if(t==='auction'){
    ch.push(p(`Owner will pay Broker a commission of ${sp} of the gross sales price excluding the buyer's premium. The commission is earned when Owner sells, exchanges, agrees to sell, or agrees to exchange all or part of the Property to anyone at any price on any terms, or when: (1) Broker individually or in cooperation with another broker procures a buyer ready, willing, and able to buy all or part of the Property at the Hidden Reserve Price or at any other price acceptable to Owner; (2) Owner grants or agrees to grant to another person an option to purchase all or part of the Property; (3) Owner transfers or agrees to transfer all or part of Owner's interest in the Property; or (4) Owner breaches this listing.`));
    ch.push(p(`Once earned, the commission is payable at the earlier of: (1) the closing and funding of any sale or exchange of all or part of the Property; (2) Owner's refusal to sell the Property after Broker's Fee has been earned; (3) Owner's breach of this Listing; or (4) at such time as otherwise set forth in this Listing.`));
  } else if(t==='sell'){
    ch.push(p(`For a sale, Owner will pay Broker a commission of ${sp} of the gross sales price, payable at closing from the sale proceeds. Owner agrees to sell the Property for the Listing Price or any other price acceptable to Owner. A commission shall be earned if Broker brings Seller an offer with terms acceptable to Seller.`));
  } else {
    ch.push(p(`If the Owner leases the Property, in whole or in part, they agree to pay the Broker a commission of ${lp} of the ${bas} rent, payable upon lease execution and tenant's commencement of rent payments.`));
    ch.push(p('The Broker is entitled to a commission under any of the following conditions:'));
    ['The Broker sells or leases the Property, or finds a buyer or tenant, based on the terms of this Agreement or other mutually acceptable terms.','The Owner sells or leases the Property directly or through another party during the term of this Agreement.',"The Owner sells or leases the Property within six (6) months after this Agreement expires to a person or entity introduced by the Broker or on Broker's registration list (including affiliates with more than a 10% ownership interest).",
     "The Property becomes unrentable or unsellable due to the Owner's voluntary or negligent actions.",'The Owner breaches this Agreement or takes any action that prevents the Broker from leasing or selling the Property.'].forEach(text=>ch.push(new Paragraph({numbering:{reference:'numbers',level:0},children:[tx(text,{size:18})]})));
    ch.push(p(`The Owner agrees to pay the Broker a commission of ${rp} of the aggregate ${bas} rent for any renewal, extension, holdover, or expansion of the lease, due at the beginning of each new term.`));
    ch.push(p(`Additionally, if the tenant or an affiliate purchases the Property during the initial lease term and within one hundred and twenty (120) days after the expiration of the lease or any renewal, extension, or expansion, the Owner must pay the Broker a commission of ${sp} of the gross selling price. If this happens, any unearned commission already received for the cancelled lease term will be credited toward the commission for the sale.`));
    if(t==='sell_lease') ch.push(p(`For a sale, Owner will pay Broker a commission of ${sp} of the gross sales price.`));
  }
  ch.push(blk());

  // ACCESS & SIGNAGE
  ch.push(sectionHdr('Access & Signage Authorization')); ch.push(blk());
  const accessPartA = s.access_signage?.part_a || 'Authorizing access to the Property means giving permission to another person to enter the Property, disclosing security codes necessary to enter the Property to such person, and lending a key to the Property to such person. To facilitate the showing and sale of the Property, Owner instructs Broker and Broker\'s associates to: (i) access the Property at reasonable times; (ii) authorize other brokers, inspectors, appraisers, lenders, engineers, surveyors, and repair persons to enter the Property at reasonable times; and (iii) duplicate keys to facilitate convenient and efficient showings.';
  ch.push(p('A.    ' + accessPartA));
  ch.push(blk());
  const accessPartB = s.access_signage?.part_b || 'Owner authorizes Broker to use reasonable efforts and act diligently to market the Property for sale, procure a buyer, and negotiate the sale of the Property.  Broker, at its direction, may cooperate with other brokers and their agents to show the Property to a prospective buyer or tenant. Owner authorizes Broker, at its discretion to offer cooperation to other brokers and their agents. Owner\'s authorization to market the Property includes the right to reproduce and distribute materials including drawings, photographs and brochures, list the Property on all forms of media and all digital platforms, to place signs and all other means. Owner will reimburse Broker for marketing per the attached Listing Agreement Addendum.';
  ch.push(p('B.    ' + accessPartB));
  ch.push(blk());
  ch.push(p('Owner directs Broker to carry out one or more of the following authorized signage installations:', {bold:true, before:0, after:40}));
  if (signs.length === 0) {
    ch.push(p('No signage authorized by Owner.', {before:0, after:60, color:'888888'}));
  }
  if (d.signBldg) {
    const signBldgText = s.signage_types?.building || 'Install Commercial Brokerage Sign(s) on building. Broker and its contractor(s) have permission to attach the sign(s) to the building using materials selected by them. Owner understands that the work may include small holes that will be drilled into the building to hold the installation pins, and that when the sign(s) are uninstalled, those pins will be removed and the holes will be sealed with silicone.';
    ch.push(p(signBldgText, {before:0, after:40}));
  }
  if (d.signAsph) {
    const signAsphText = s.signage_types?.asphalt || 'Install Commercial Brokerage Sign(s) anchored to asphalt using rebar spikes. Broker and its contractor(s) have permission to attach the sign(s) to the asphalt lot using materials selected by them. Owner understands that the work may include 12 small holes (4 per post) that will be drilled into the asphalt to hold the rebar spikes, and when the sign(s) are uninstalled, those spikes will be removed and the holes can be filled with silicone upon Owner\'s request.';
    ch.push(p(signAsphText, {before:0, after:40}));
  }
  if (d.signFnce) {
    const signFnceText = s.signage_types?.fence || 'Install Commercial Sign(s) on Fence. Broker and its contractor(s) have permission to attach the sign(s) to the fence using materials they select. Owner understands this may result in wear/scratching/minor impact on the fence.';
    ch.push(p(signFnceText, {before:0, after:40}));
  }
  if (d.signWndw) {
    const signWndwText = s.signage_types?.window || 'Install Commercial Window Sign(s). Broker and its contractor(s) have permission to attach the window sign(s) to the building using materials selected by them. Broker does not recommend window signs to be installed on windows that are tinted. If a window sign is installed, residue may be left from the sign, burning may occur in the tint, and/or ripping of tint may occur.';
    ch.push(p(signWndwText, {before:0, after:40}));
  }
  if (d.signYard) {
    const signYardText = s.signage_types?.yard || 'Install Commercial Yard Sign(s). Broker and its contractor(s) have permission to install a yard sign.';
    ch.push(p(signYardText, {before:0, after:40}));
  }
  ch.push(blk());

  // ADDITIONAL TERMS
  ch.push(sectionHdr('Additional Terms & Conditions')); ch.push(blk());

  ch.push(legalHdr('Protection Period'));
  const protectionPeriodText = s.type_specific?.[t]?.protection_period || (t==='sell' ? 'A "Protection Period" is a period of one hundred eighty (180) days commencing on the day immediately following the expiration or termination of this Listing Agreement. No later than thirty (30) days after the termination of this Listing Agreement, the Broker must provide the Owner with a written notice containing the names of all individuals or entities to whom the Broker introduced the Property during the term of this Listing.\n\nIf the Owner leases, sells, or otherwise transfers all or any part of the Property to any individual or entity named in the Broker\'s notice, or to an affiliate or entity controlled by such individual or entity, during the Protection Period, the Owner shall pay the Broker the commission that would have been due had the transaction occurred while this Listing Agreement was in effect. This clause shall survive the termination or expiration of this Agreement.' : 'A "Protection Period" is a six-month period commencing on the day immediately following the expiration or termination of this Listing Agreement. No later than thirty (30) days after the termination of this Listing Agreement, the Broker must provide the Owner with a written notice containing the names of all individuals or entities to whom the Broker introduced the Property during the term of this Listing.\n\nIf the Owner leases, sells, or otherwise transfers all or any part of the Property to any individual or entity named in the Broker\'s notice, or to a relative or business associate of such individual or entity, during the Protection Period, the Owner shall pay the Broker the commission that would have been due had the transaction occurred while this Listing Agreement was in effect. This clause shall survive the termination or expiration of this Agreement.');
  protectionPeriodText.split('\n\n').forEach(paragraph => ch.push(p(paragraph)));
  ch.push(blk());

  ch.push(legalHdr('Owner Representations & Warranties'));
  const repWarrantiesText = s.type_specific?.[t]?.rep_and_warranties || `Except as provided otherwise in this Listing, Owner represents and warrants that: (a) Owner holds fee simple title to the Property and has the legal right to enter into this Agreement; (b) Owner is not bound by a listing agreement with another broker for the sale, exchange, or lease of the Property that is or will be in effect during the Term of this Agreement; (c) no person or entity has any right to purchase, lease, or acquire the Property by an option, right of refusal, or other agreement; (d) there are no delinquencies or defaults under any mortgage or other encumbrance on the Property; (e) the Property is not subject to the jurisdiction of any court; (f) Owner owns sufficient intellectual property rights in any materials which Owner provides to Broker related to the Property (for example, brochures, photographs, drawings, or articles) to permit Broker to reproduce and distribute the materials in marketing the Property or for other purposes related to this Agreement; (g) Owner has reviewed this Agreement and all information relating to the Property which Owner provides to Broker is true and correct to the best of Owner's knowledge; and (h) the signers below all constitute all parties required to execute this Agreement and have full authority to execute this Agreement.`;
  ch.push(p(repWarrantiesText));
  ch.push(blk());

  ch.push(legalHdr('Indemnity'));
  const indemnityText = s.common?.indemnity || `Owner recognizes that the Broker, cooperating broker and agents ("Brokers") involved in the sale and/or lease are relying on all information provided or supplied by Owner or Owner's sources and/or Buyer or Buyer's sources and/or Tenant or Tenant's sources in connection with the Property. Owner agrees to indemnify, defend and hold harmless the Brokers, their agents and employees from any claims, demands, suits, liabilities, costs and expenses (including reasonable attorney's fees) arising out of any misrepresentation or concealment of facts by Owner or Owner's sources and/or Buyer or Buyer's sources.`;
  ch.push(p(indemnityText));
  ch.push(blk());

  ch.push(legalHdr('Default'));
  const defaultText = s.common?.default || "If Owner breaches this Agreement, Owner is in default and will be liable to Broker for the amount of Broker's fee specified in this Agreement and any other fees Broker is entitled to receive under this Listing along with all costs expended by Broker and its agents in connection with the listing, including costs of marketing, advertising, signage, mileage and the like; Broker may also terminate this Listing and exercise any other remedy at law or in equity. If a sales price is not determinable in the event of an exchange or breach of this Agreement, the Listing Price will be the sales price for the purpose of calculating Broker's fee. Interest is calculated on amounts due at 10% / annum or the highest legal rate whichever is less.";
  ch.push(p(defaultText));
  ch.push(blk());

  ch.push(legalHdr('Mediation'));
  const mediationText = s.common?.mediation || "The parties agree to negotiate in good faith in an effort to resolve any dispute that may arise between the parties. If the dispute cannot be resolved by negotiation, the parties will submit the dispute to mediation. The parties to the dispute will choose a mutually acceptable mediator and will share the costs of mediation equally.";
  ch.push(p(mediationText));
  ch.push(blk());

  ch.push(legalHdr("Attorney's Fees"));
  const feesText = s.common?.attorneys_fees || "If Owner or Broker is a prevailing party in any legal proceeding brought as a result of a dispute under this Agreement or any transaction related to or contemplated by this Agreement, such party may recover from the non-prevailing party all costs of such proceeding and reasonable attorney's fees.";
  ch.push(p(feesText));
  ch.push(blk());

  ch.push(legalHdr('Notices'));
  const noticesText = s.common?.notices || "Unless specifically stated otherwise in this Agreement, all notices, waivers, and demands required or permitted hereunder shall be in writing and delivered to the addresses first set forth above, by one of the following methods: (a) hand delivery, whereby delivery is deemed to have occurred at the time of delivery; (b) a nationally recognized overnight courier company, whereby delivery is deemed to have occurred on delivery; or (c) email provided that the transmission is completed no later than 4:00 p.m. Eastern Time and a confirmation is received, whereby delivery is deemed to have occurred on the day on which electronic transmission is completed.";
  ch.push(p(noticesText));
  ch.push(blk());

  ch.push(legalHdr('Counterparts'));
  const counterpartsText = s.common?.counterparts || "This Agreement may be executed by the Parties in separate counterparts, each of which when so executed and delivered shall be an original for all purposes, but all such counterparts shall together constitute but one and the same instrument. A signed copy of this Agreement shall be delivered by email shall be deemed to have the same legal effect as delivery of an original signed copy of this Agreement.";
  ch.push(p(counterpartsText));
  ch.push(blk());

  ch.push(legalHdr('Severability / Sole Agreement'));
  const severabilityText = s.common?.severability || "If any provision of this Agreement is determined to be unenforceable, the remaining terms and provisions shall not in any way be construed, impaired or invalidated. This constitutes the entire agreement between Owner and Broker. No oral or implied agreement or understanding shall cancel or vary the Agreement terms. Any amendments shall be made in writing, signed by both parties. This Agreement is binding upon the parties, their heirs, administrators, executors, successors, and assigns.";
  ch.push(p(severabilityText));
  ch.push(blk());

  ch.push(legalHdr('Dual Agency'));
  const dualAgencyText = s.common?.dual_agency || "Owner acknowledges receipt of any forms, disclosures, statistics, and policies, if any and if applicable as mandated by state law. Broker may act as a dual agent by representing both Owner and Buyer/Tenant in a transaction contemplated by this Agreement only if both parties to the transaction consent after having been informed of the dual agency relationship. Broker shall not permit another salesperson or agent affiliated with Broker to represent another party in a transaction involving the Broker (whether as the exclusive agent for that party, a subagent or dual agent) without obtaining the written consent of both parties to the transaction. Should a dual agency relationship arise, Owner and Buyer will be provided with a dual agency disclosure form describing the agent's duties and Owner's and Buyer's options if they elect not to consent to the dual agency relationship.";
  ch.push(p(dualAgencyText));
  ch.push(blk());

  ch.push(legalHdr('Disclosure and Owner\'s Covenants'));
  const disclosureText = s.common?.disclosure || "Owner specifically acknowledges and understands that if Owner knows, or acquires knowledge after the execution of this Listing Agreement, of facts, environmental or otherwise, adversely affecting or materially impairing the title, condition, compliance with applicable laws, value or desirability of the Property, whether such facts are readily observable or not, then Owner shall disclose such facts to Broker and Buyer/Tenant. If Owner knows of such facts, Owner shall set them forth by written document attached to this Agreement. Owner has fully reviewed this Agreement and the document(s) attached, if any, relating to the Property, and Owner represents, to the best of its knowledge, that such information is true and accurate. Owner's obligation to indemnify, set out in this Agreement shall apply to any claim arising out of or resulting from the inaccuracy of information and from Owner's failure to disclose any facts, environmental or otherwise, adversely affecting or materially impairing the value or desirability of the Real Estate.";
  ch.push(p(disclosureText));
  ch.push(blk());

  ch.push(legalHdr("Owner's Acknowledgments"));
  const ackA = s.common?.owners_acknowledgments_a || "Should there be any tax, fee or other cost required in regard to the payment due Broker, such tax, fee or other cost shall be paid by Owner.";
  const ackB = s.common?.owners_acknowledgments_b || "Owner acknowledges that there are currently no other listing contracts to sell, exchange or lease the Property.";
  const ackC = s.common?.owners_acknowledgments_c || "Owner acknowledges and agrees to comply with all applicable federal, state and local laws, regulations, codes, ordinances and administrative orders having jurisdiction over the parties, the Property or the subject matters of this Agreement, including, but not limited to, the 1964 Civil Rights Act and all amendments thereto, the Foreign Investment in Real Property Tax Act, the Comprehensive Environmental Response Compensation and Liability Act, and the Americans with Disabilities Act.";
  const ackD = s.common?.owners_acknowledgments_d || "Broker is not an expert in legal matters or other areas such as finance, property inspection, environmental and the like. Broker should consult qualified professionals and resources including a lawyer. Should Broker provide referrals or recommend resources, Owners agrees that Broker neither warrants or guarantees or endorses those entities or persons.";
  ch.push(p("a. " + ackA));
  ch.push(p("b. " + ackB));
  ch.push(p("c. " + ackC));
  ch.push(p("d. " + ackD));
  ch.push(blk());
  ch.push(p("Questions regarding this Agreement, its Addendum, attachments and accompanying disclosure forms or with regard to obligations in a purchase contract should be directed to Owner's lawyer.", {italics:true, color:'444444'}));
  ch.push(blk());

  if(d.mktgReimb){ ch.push(legalHdr('Marketing Reimbursement')); ch.push(p(`Owner will reimburse Brokerage for marketing in the amount of ${d.mktgReimb} in the event of a termination by Owner prior to the end of the terms of the Agreement.`)); ch.push(blk()); }
  if(d.addTerms){ ch.push(legalHdr('Additional Terms')); d.addTerms.split('\n').filter(Boolean).forEach(line=>ch.push(p(line))); ch.push(blk()); }

  // SIGNATURES
  ch.push(sectionHdr('Signatures')); ch.push(blk());
  ch.push(new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[half,half], rows:[new TableRow({ children:[
    new TableCell({ borders:allBdrs(), margins:{top:120,bottom:200,left:160,right:160}, width:{size:half,type:WidthType.DXA}, children:[
      new Paragraph({children:[tx(ownerLbl.toUpperCase(),{bold:true,color:C.PINK,size:17})]}),
      ...(d.ownerName?[new Paragraph({spacing:{before:80,after:80},children:[tx(d.ownerName,{size:18})]})]:[]),
      ...sigLine('Signature'),...sigLine('Printed Name'),...sigLine('Title'),...sigLine('Date'),
    ]}),
    new TableCell({ borders:allBdrs(), margins:{top:120,bottom:200,left:160,right:160}, width:{size:half,type:WidthType.DXA}, children:[
      new Paragraph({children:[tx('BROKER — ASTONISH LLC',{bold:true,color:C.PINK,size:17})]}),
      new Paragraph({spacing:{before:80,after:80},children:[tx('Astonish LLC DBA Astonish Commercial Real Estate Services',{size:17})]}),
      ...sigLine('Signature'),
      new Paragraph({spacing:{before:80,after:40},children:[tx(`Printed Name: ${d.agentName||(s.defaults?.agent_name||'Michael Bergman')}`,{size:18})]}),
      new Paragraph({spacing:{before:0,after:40},children:[tx(`Title: ${d.agentTitle||(s.defaults?.agent_title||'Broker')}`,{size:18})]}),
      ...(d.agentLicense?[new Paragraph({spacing:{before:0,after:80},children:[tx(`License No.: ${d.agentLicense}`,{size:18})]})]:[]),
      ...sigLine('Date'),
    ]}),
  ]})]  }));
  ch.push(blk());
  ch.push(new Paragraph({spacing:{before:80,after:0},alignment:AlignmentType.CENTER,children:[tx("Questions regarding this Agreement, its Addendum, attachments and accompanying disclosure forms or with regard to obligations in any contract should be directed to Owner's lawyer.",{size:16,color:'888888',italics:true})]}));

  return new Document({
    numbering:{config:[{reference:'numbers',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
    styles:{default:{document:{run:{font:'Arial',size:18}}}},
    sections:[{
      properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
      headers:{default:new Header({children:[
        new Table({width:{size:10080,type:WidthType.DXA},columnWidths:[5040,5040],rows:[new TableRow({children:[
          new TableCell({borders:noBdrs(),width:{size:5040,type:WidthType.DXA},children:[new Paragraph({children:[tx('Astonish',{bold:true,size:32,font:'Georgia',color:C.NAVY}),tx('  Commercial Real Estate Services',{size:16,color:'888888'})]})] }),
          new TableCell({borders:noBdrs(),width:{size:5040,type:WidthType.DXA},children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[tx(`Listing Agreement — ${LABELS[t]}`,{size:15,color:'888888'})]})] }),
        ]})]})  ,
        new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:8,color:C.PINK}},spacing:{after:0},children:[tx('')]}),
      ]})},
      footers:{default:new Footer({children:[
        new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:C.CYAN}},spacing:{before:80},children:[
          tx(`Astonish LLC  ·  ${agentAddr||(s.defaults?.agent_address||'9918 Carver Rd., Suite 101, Cincinnati OH 45242')}  ·  ${d.agentPhone||(s.defaults?.agent_phone||'513.334.3624')}  ·  ${d.agentEmail||(s.defaults?.agent_email||'info@astonishcommercial.com')}`,{size:14,color:'888888'}),
          tx('     Page ',{size:14,color:'888888'}),
          new TextRun({children:[PageNumber.CURRENT],size:14,font:'Arial',color:'888888'}),
        ]}),
        new Paragraph({children:[tx(`License No. ${d.agentLicense||(s.defaults?.agent_license||'BRKM.2020007859')}  ·  Generated: ${today}`,{size:13,color:'AAAAAA'})]}),
      ]})},
      children:ch,
    }],
  });
}
