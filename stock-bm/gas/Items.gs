/**
 * รายการยาของคลังใน รพ.สต.บ้านบวกหมื้อ — ดึงจาก DrugMaster ของระบบคลังยาใหญ่ (ไม่มีรายการฝังในโค้ด)
 *
 * เพิ่มยาใหม่ภายหลัง ทำได้ 3 ทาง
 *   1) เพิ่มยาใน DrugMaster ของระบบใหญ่ตามปกติ แล้วกดเมนู คลังใน → ดึงยาใหม่จากคลังยาใหญ่
 *      (ระบบดึงให้เองทุกคืนด้วย ถ้ากดรัน installHubSync() ไว้หนึ่งครั้ง)
 *   2) เมนู คลังใน → เพิ่มยาใหม่ + สร้าง QR
 *   3) พิมพ์ชื่อยาในแถวว่างท้ายแท็บ "รายการยา" ระบบใส่รหัสให้เอง
 * ยาที่มีอยู่แล้วจะไม่ถูกแก้ชื่อ หมวด จุดสั่งซื้อ หรือติ๊กใช้งาน ระบบอัปเดตให้แค่ช่อง "ทุน/หน่วย"
 * รหัส B0001, B0002, … ไม่เปลี่ยนเมื่อดึงซ้ำ สติกเกอร์ QR ที่พิมพ์แล้วจึงใช้ได้ตลอด
 */
const ITEM_HEADER = ['รหัส', 'ชื่อยา', 'ชื่อไทย', 'หน่วยนับ', 'หมวด/ชั้น', 'จุดสั่งซื้อ', 'ใช้งาน', 'ชื่อในคลังบวกหมื้อ', 'ทุน/หน่วย'];
const CODE_PREFIX = 'B', CODE_DIGITS = 4;

/* หมวดแบบเดียวกับคลังในบ้านเป้า 8 หมวด
   1 ยาฉีด · 2 ยาใช้ภายนอก · 3 ยาน้ำ · 4 ยาเม็ด · 5 สมุนไพร · 6 ยาโรคเรื้อรัง · 7 เวชภัณฑ์ · 8 วัสดุ Lab
   ลำดับ: ชื่อตรงกับรายการบ้านเป้า → ใช้หมวดเดียวกับบ้านเป้า · ไม่ตรง → จัดตามชนิดยาจากชื่อ (และหน่วย) */
const SHELVES = ['1 ยาฉีด', '2 ยาใช้ภายนอก', '3 ยาน้ำ', '4 ยาเม็ด', '5 สมุนไพร', '6 ยาโรคเรื้อรัง', '7 เวชภัณฑ์', '8 วัสดุ Lab'];
const BANPAO_SHELF = {
  "1243a steri gage steam integrator 3m 500 ea": "7 เวชภัณฑ์",
  "1292 steam rapid attest 3m tube": "7 เวชภัณฑ์",
  "acyclovir tab 400 mg": "4 ยาเม็ด",
  "adrenaline inj 1 mg/ml": "1 ยาฉีด",
  "albendazole susp 100 mg/5 ml (20 ml)": "3 ยาน้ำ",
  "albendazole tab 200 mg": "4 ยาเม็ด",
  "allopurinol tab 100 mg": "6 ยาโรคเรื้อรัง",
  "amiloride hydrochloride + hydrochlorothiazide tab 5 mg +50 mg": "6 ยาโรคเรื้อรัง",
  "amitriptyline tab 10 mg": "4 ยาเม็ด",
  "amlodipine tab 5 mg": "6 ยาโรคเรื้อรัง",
  "amoxicillin cap 500 mg": "4 ยาเม็ด",
  "amoxicillin syr 125 mg/5 ml (60": "3 ยาน้ำ",
  "amoxicillin/clavulanic acid syr 457 mg/5 ml (70": "3 ยาน้ำ",
  "amoxicillin/clavulanic acid tab 875/125 mg": "4 ยาเม็ด",
  "antacid syr 240 ml": "3 ยาน้ำ",
  "antazoline hydrochloride + tetryzoline hydrochlori eye (0.5 mg+0.4 mg)": "2 ยาใช้ภายนอก",
  "aromatic ammonia spi 450 ml": "2 ยาใช้ภายนอก",
  "aspirin ec 81 mg": "6 ยาโรคเรื้อรัง",
  "atenolol tab 50 mg": "6 ยาโรคเรื้อรัง",
  "atorvastatin tab 40 mg": "6 ยาโรคเรื้อรัง",
  "autoclave rol 0.5\"": "7 เวชภัณฑ์",
  "benzyl benzoate lotion 25% (60 ml)": "2 ยาใช้ภายนอก",
  "betahistine mesilate tab 6 mg": "4 ยาเม็ด",
  "betamethasome cre 5 g": "2 ยาใช้ภายนอก",
  "blade no. 11": "7 เวชภัณฑ์",
  "blade no. 15": "7 เวชภัณฑ์",
  "bottle ขวดพลาสติก แบน 30 ซีซี": "7 เวชภัณฑ์",
  "brompheniramine + phenylephrine syr (2 mg +5 mg)/5": "3 ยาน้ำ",
  "brompheniramine + phenylephrine tab 4 mg +10 mg": "4 ยาเม็ด",
  "brown mixture gpo mix 60 ml": "3 ยาน้ำ",
  "budesonide mdi 200 mcg": "2 ยาใช้ภายนอก",
  "calamine lot 60 ml": "2 ยาใช้ภายนอก",
  "calcium carbonate tab 1500 mg": "4 ยาเม็ด",
  "carminative mixture mix 180 ml": "3 ยาน้ำ",
  "cetirizine syr 1 mg/ml (60 ml)": "3 ยาน้ำ",
  "cetirizine tab 10 mg": "4 ยาเม็ด",
  "chlorhexidine scrub sol 450 ml": "2 ยาใช้ภายนอก",
  "chlorpheniramine inj 10 mg/ml": "1 ยาฉีด",
  "chlorpheniramine syr 2 mg/5 ml (60 m": "3 ยาน้ำ",
  "chlorpheniramine tab 4 mg": "4 ยาเม็ด",
  "clindamycin cap 300 mg": "4 ยาเม็ด",
  "clotrimazole cre 1% (15 g)": "2 ยาใช้ภายนอก",
  "clotrimazole vag 500 mg": "2 ยาใช้ภายนอก",
  "colchicine tab 0.6 mg": "6 ยาโรคเรื้อรัง",
  "condom mm 49": "7 เวชภัณฑ์",
  "condom mm 52": "7 เวชภัณฑ์",
  "cotton g 0.35": "7 เวชภัณฑ์",
  "cotton stick no. l": "7 เวชภัณฑ์",
  "cotton stick no. s": "7 เวชภัณฑ์",
  "dexamethasone + neomycin eye 4 ml": "2 ยาใช้ภายนอก",
  "dexamethasone inj 4mg/ml": "1 ยาฉีด",
  "dextose 50% inj 50 ml": "1 ยาฉีด",
  "dextromethorphan tab 15 mg": "4 ยาเม็ด",
  "diazepam tab 5 mg": "4 ยาเม็ด",
  "diclofenac tab 25 mg": "4 ยาเม็ด",
  "dicloxacillin cap 250 mg": "4 ยาเม็ด",
  "dicloxacillin syr 62.5 mg/5ml": "3 ยาน้ำ",
  "digoxin tab 0.25 mg": "6 ยาโรคเรื้อรัง",
  "dimenhydrinate inj 50 mg/ml": "1 ยาฉีด",
  "dimenhydrinate tab 50 mg": "4 ยาเม็ด",
  "diphtheria and tetanus vaccine(dt) inj 0.5 ml/dose (0.": "1 ยาฉีด",
  "disp. glove no. m": "7 เวชภัณฑ์",
  "disp. glove no. s": "7 เวชภัณฑ์",
  "disp. mask pcs 50": "7 เวชภัณฑ์",
  "disp. needle no. 18 x 1.5\"": "7 เวชภัณฑ์",
  "disp. needle no. 21 x 1.5\"": "7 เวชภัณฑ์",
  "disp. needle no. 23 x 1.5\"": "7 เวชภัณฑ์",
  "disp. needle no. 24 x 1.5\"": "7 เวชภัณฑ์",
  "disp. needle no. 25 x 1\"": "7 เวชภัณฑ์",
  "disp. syringe insulin ติดเข็มถอดหัวได้ ml 1(27gx0.5\")": "7 เวชภัณฑ์",
  "disp. syringe insulin ติดเข็มถอดหัวไม่ได้ ml 1(29g x 0.5\")": "7 เวชภัณฑ์",
  "disp. syringe ml 10": "7 เวชภัณฑ์",
  "disp. syringe ml 3": "7 เวชภัณฑ์",
  "disp. syringe ml 5": "7 เวชภัณฑ์",
  "domperidone syr 5 mg/5 ml": "3 ยาน้ำ",
  "domperidone tab 10 mg": "4 ยาเม็ด",
  "doxazosin tab 2 mg": "6 ยาโรคเรื้อรัง",
  "elastic bandage no. 3\"": "7 เวชภัณฑ์",
  "elastic bandage no. 4\"": "7 เวชภัณฑ์",
  "elastic bandage no. 6\"": "7 เวชภัณฑ์",
  "enalapril tab 20 mg": "6 ยาโรคเรื้อรัง",
  "enalapril tab 5 mg": "6 ยาโรคเรื้อรัง",
  "erythromycin syr 125 mg/5 ml": "3 ยาน้ำ",
  "ethanol sol 70% 450 ml": "2 ยาใช้ภายนอก",
  "ethinylestradiol + levonorgestrel แผง 30 mcg+ 0.15 mg": "4 ยาเม็ด",
  "fenofibrate micronized tab 100 mg": "6 ยาโรคเรื้อรัง",
  "ferrous fumarate drop susp 15 mg/0.6 ml": "3 ยาน้ำ",
  "ferrous fumarate tab 200 mg": "6 ยาโรคเรื้อรัง",
  "fluconazole cap 200 mg": "4 ยาเม็ด",
  "foley catheter no. 14": "7 เวชภัณฑ์",
  "foley catheter no. 16": "7 เวชภัณฑ์",
  "foley catheter no. 18": "7 เวชภัณฑ์",
  "folic acid tab 5 mg": "6 ยาโรคเรื้อรัง",
  "furosemide tab 40 mg": "6 ยาโรคเรื้อรัง",
  "gauze bandages หลา 3\" x 6": "7 เวชภัณฑ์",
  "gauze bandages หลา 4\" x 6": "7 เวชภัณฑ์",
  "gauze bandages หลา 6\" x 6": "7 เวชภัณฑ์",
  "gemfibrozil cap 600 mg": "6 ยาโรคเรื้อรัง",
  "glipizide tab 5 mg": "6 ยาโรคเรื้อรัง",
  "guaifenesin syr 100 mg/5 ml (60": "3 ยาน้ำ",
  "guaifenesin tab 100 mg": "4 ยาเม็ด",
  "guaze pad no. 3 x 4\"": "7 เวชภัณฑ์",
  "hydralazine tab 25 mg": "6 ยาโรคเรื้อรัง",
  "hydrochlorothiazide tab 50 mg": "6 ยาโรคเรื้อรัง",
  "hydrogen peroxide sol 450 ml": "2 ยาใช้ภายนอก",
  "hydroxyzine tab 10 mg": "4 ยาเม็ด",
  "hyoscine inj 20 mg/ml": "1 ยาฉีด",
  "hyoscine syr 5 mg/5 ml": "3 ยาน้ำ",
  "hyoscine tab 10 mg": "4 ยาเม็ด",
  "ibuprofen susp 100 mg/5 ml (60 ml)": "3 ยาน้ำ",
  "ibuprofen tab 400 mg": "4 ยาเม็ด",
  "injection plug pcs supply": "7 เวชภัณฑ์",
  "insulin 70/30 inj 10 ml": "1 ยาฉีด",
  "insulin nph inj vial 100 iu/ml (10 ml)": "1 ยาฉีด",
  "ipratropium + fenoterol mdi 20 mcg+50 mcg (": "2 ยาใช้ภายนอก",
  "ipratropium + fenoterol neb (0.5 mg+1.25 mg": "2 ยาใช้ภายนอก",
  "isosorbide dinitrate sl tab 5 mg": "6 ยาโรคเรื้อรัง",
  "isosorbide dinitrate tab 10 mg": "6 ยาโรคเรื้อรัง",
  "iv catheter no. 22": "7 เวชภัณฑ์",
  "iv catheter no. 24": "7 เวชภัณฑ์",
  "iv set pcs supply": "7 เวชภัณฑ์",
  "k-y gelly g 82": "7 เวชภัณฑ์",
  "lidocaine inj 2% (20 ml)": "1 ยาฉีด",
  "lithium battery cr2032 1 pack": "8 วัสดุ Lab",
  "lorazepam tab 1 mg": "4 ยาเม็ด",
  "losartan tab 50 mg": "6 ยาโรคเรื้อรัง",
  "manidipine tab 20 mg": "6 ยาโรคเรื้อรัง",
  "medroxyprogesterone inj 150 mg/3 ml": "1 ยาฉีด",
  "metformin tab 500 mg": "6 ยาโรคเรื้อรัง",
  "methyl salicylate cream, compound (ครีมน้ำมันระกำ) cre 15 g": "2 ยาใช้ภายนอก",
  "methyldopa tab 250 mg": "6 ยาโรคเรื้อรัง",
  "metoclopramide inj 10 mg/2 ml": "1 ยาฉีด",
  "metoprolol tab 100 mg": "6 ยาโรคเรื้อรัง",
  "metronidazole tab 400 mg": "4 ยาเม็ด",
  "milk of magnesia sus 240 ml": "3 ยาน้ำ",
  "multivitamin tab": "4 ยาเม็ด",
  "neofix (carifix stretch) no. 5 cm x 10 m": "7 เวชภัณฑ์",
  "neomycin sulfate + polymyxin b sulfate+gramicidin eye (2 mg +5,000iu": "2 ยาใช้ภายนอก",
  "ng tube no. 14": "7 เวชภัณฑ์",
  "ng tube no. 16": "7 เวชภัณฑ์",
  "norfloxacin tab 400 mg": "4 ยาเม็ด",
  "nss inj 100 ml": "1 ยาฉีด",
  "nss inj 1000 ml": "1 ยาฉีด",
  "nss inj 5 ml": "1 ยาฉีด",
  "nss irrigate sol 1000 ml": "2 ยาใช้ภายนอก",
  "nylon 3/0 24 mm 3/8 round": "7 เวชภัณฑ์",
  "nylon 4/0 30 mm 3/8 round": "7 เวชภัณฑ์",
  "omeprazole cap 20 mg": "4 ยาเม็ด",
  "omeprazole inj 40 mg": "1 ยาฉีด",
  "orphenadrine + paracetamol tab 35 mg +500 mg": "4 ยาเม็ด",
  "ors adult sac": "4 ยาเม็ด",
  "ors ped sac": "4 ยาเม็ด",
  "oxygen cannular no. adult": "7 เวชภัณฑ์",
  "oxygen mask with bag no. adult": "7 เวชภัณฑ์",
  "oxygen nebulizer no. adult": "7 เวชภัณฑ์",
  "oxygen nebulizer no. child": "7 เวชภัณฑ์",
  "paracetamol dro 60 mg/0.6 ml": "3 ยาน้ำ",
  "paracetamol syr 120 ml/5 mg": "3 ยาน้ำ",
  "paracetamol tab 325 mg": "4 ยาเม็ด",
  "paracetamol tab 500 mg": "4 ยาเม็ด",
  "phenytoin cap 100 mg": "6 ยาโรคเรื้อรัง",
  "pioglitasone tab 30 mg": "6 ยาโรคเรื้อรัง",
  "plaster micropore no. 0.5\"": "7 เวชภัณฑ์",
  "plaster micropore no. 1\"": "7 เวชภัณฑ์",
  "plaster transpore no. 0.5\"": "7 เวชภัณฑ์",
  "plaster transpore no. 1\"": "7 เวชภัณฑ์",
  "povidine scrub soln 7.5% (450 ml)": "2 ยาใช้ภายนอก",
  "povidine sol 10% (450 ml)": "2 ยาใช้ภายนอก",
  "prednisolone tab 5 mg": "4 ยาเม็ด",
  "pregnancy test (hcg urine strip) 100 test": "8 วัสดุ Lab",
  "propanolol tab 10 mg": "6 ยาโรคเรื้อรัง",
  "rabies vaccine inj": "1 ยาฉีด",
  "salbutamol mdi 0.1 mg/dose": "2 ยาใช้ภายนอก",
  "salbutamol neb 2.5 mg/2.5 ml": "2 ยาใช้ภายนอก",
  "salbutamol syr 2 mg/5 ml (60 m": "3 ยาน้ำ",
  "salbutamol tab 2 mg": "6 ยาโรคเรื้อรัง",
  "salmeterol + fluticasone evo 25/125": "2 ยาใช้ภายนอก",
  "silk ไหมม้วนเย็บแผล no. 3/0": "7 เวชภัณฑ์",
  "silver sulfadiazine cre 25 g": "2 ยาใช้ภายนอก",
  "simethicone tab 80 mg": "4 ยาเม็ด",
  "simvastatin tab 20 mg": "6 ยาโรคเรื้อรัง",
  "sofra tulle no. 10 cm x 10 cm": "7 เวชภัณฑ์",
  "spironolactone tab 25 mg": "6 ยาโรคเรื้อรัง",
  "sterile glove no. 6.5": "7 เวชภัณฑ์",
  "sterile glove no. 7": "7 เวชภัณฑ์",
  "sterile water for injection inj 10 ml": "1 ยาฉีด",
  "strips (แถบตรวจน้ำตาล) box 100 pcs": "8 วัสดุ Lab",
  "sulfamethoxazole + trimethoprim (co-trimoxazole) syr 200mg+40mg/5ml": "3 ยาน้ำ",
  "ta 0.02% cre 5 g": "2 ยาใช้ภายนอก",
  "ta 0.1% cre 5 g": "2 ยาใช้ภายนอก",
  "theophylline tab 200 mg": "6 ยาโรคเรื้อรัง",
  "top dressing (sterile) no. 4\" x 6\"": "7 เวชภัณฑ์",
  "tramadol cap 50 mg": "4 ยาเม็ด",
  "triamcinolone oral paste sac 1 g": "2 ยาใช้ภายนอก",
  "triferdine tab 150 mg": "4 ยาเม็ด",
  "urine bag เทล่าง pcs supply": "7 เวชภัณฑ์",
  "urine strip 2 gp 100 test": "8 วัสดุ Lab",
  "vitamin b complex tab": "4 ยาเม็ด",
  "vitamin b1 tab 100 mg": "4 ยาเม็ด",
  "vitamin c tab 100 mg": "4 ยาเม็ด",
  "water for irrigation sol 1000 ml": "2 ยาใช้ภายนอก",
  "ขมิ้นชัน cap 500 mg": "5 สมุนไพร",
  "ครีมพญายอ สมุนไพรไทย cre 5 g": "5 สมุนไพร",
  "ชุดเครื่องตรวจวัดระดับน้ำตาลในเลือด ชุด set": "8 วัสดุ Lab",
  "ซองซิปสีชาทึบ cm 10 x 15": "7 เวชภัณฑ์",
  "ซองซิปสีชาทึบ cm 15 x 23": "7 เวชภัณฑ์",
  "ซองซิปสีชาทึบ cm 9 x 13": "7 เวชภัณฑ์",
  "ซองซิปสีใส cm 10 x 15": "7 เวชภัณฑ์",
  "ซองซิปสีใส cm 15 x 23": "7 เวชภัณฑ์",
  "ซองซิปสีใส cm 20 x 30": "7 เวชภัณฑ์",
  "ซองซิปสีใส cm 9 x 13": "7 เวชภัณฑ์",
  "ดินน้ำมัน 1 pack": "8 วัสดุ Lab",
  "น้ำยาล้างเครื่องมือแพทย์ sol 5 l (สีใส)": "7 เวชภัณฑ์",
  "ฟ้าทะลายโจร cap 500 mg": "5 สมุนไพร",
  "มะขามแขก tab 400 mg": "5 สมุนไพร",
  "ยาชงหญ้าหนวดแมว pow 2 g": "5 สมุนไพร",
  "ยาธาตุอบเชย sus 120 ml": "5 สมุนไพร",
  "ยาน้ำแก้ไอมะขามป้อม sus 120 ml": "5 สมุนไพร",
  "ยารางจืด cap 400 mg": "5 สมุนไพร",
  "ยาหอมเทพจิตร sachet 200 mg": "5 สมุนไพร",
  "ยาอมมะแว้ง รสบ๊วย tab 20 เม็ดต่อซอง": "5 สมุนไพร",
  "ยาไพล cre 30 g": "5 สมุนไพร",
  "สติ๊กเกอร์ยา ขนาด 8x9 cm rol 500 ดวง": "7 เวชภัณฑ์",
  "สหัศธาราแคปซูล cap 500 mg": "5 สมุนไพร",
  "สาย suction no. 14": "7 เวชภัณฑ์",
  "สาย suction no. 16": "7 เวชภัณฑ์",
  "หลอด hematocrit สีแดง 100 tip": "8 วัสดุ Lab",
  "เข็มเจาะปลายนิ้ว box 200 pcs": "8 วัสดุ Lab",
  "เถาวัลย์เปรียง cap 500 mg": "5 สมุนไพร",
  "เพชรสังฆาต cap 500 mg": "5 สมุนไพร"
};
const CHRONIC_ = /\b(aspirin|allopurinol|amiloride|amlodipine|atenolol|atorvastatin|carvedilol|clopidogrel|colchicine|digoxin|doxazosin|enalapril|fenofibrate|ferrous|folic|furosemide|gemfibrozil|glipizide|hydralazine|hydrochlorothiazide|isosorbide|losartan|manidipine|metformin|methyldopa|metoprolol|nifedipine|pioglit\w*|phenytoin|phenyltoin|propr?anolol|simvastatin|spironolactone|theophylline|salbutamol|levothyroxine|methimazole|propylthiouracil|alfacalcidol|sodium bicarbonate|finasteride|phenobarbitone|trihexyphenidyl|methotrexate|montelukast)\b/;
function guessShelf_(name, unit) {
  const raw = String(name || '').trim();
  const k = raw.toLowerCase().replace(/\s+/g, ' ');
  if (BANPAO_SHELF[k]) return BANPAO_SHELF[k];
  const n = ' ' + k + ' ' + String(unit || '').toLowerCase() + ' ';
  const has = re => re.test(n);
  if (has(/strips|แถบตรวจ|เข็มเจาะปลายนิ้ว|hematocrit|pregnancy|\bupt\b|hcg|ดินน้ำมัน|เครื่องตรวจน้ำตาล|ถ่าน|lithium battery|ชุดตรวจ|lancet/)) return '8 วัสดุ Lab';
  if (has(/needle|syringe|syring\b|glove|mask|blade|cotton|condom|catheter|bandage|gauze|guaze|ng tube|oxygen|plaster|sofra|dressing|neofix|nylon|silk|\bsteri\b|attest|autoclave|integrator|spore|k-y|ซองซิ[ปบ]|suction|urine bag|iv set|injection plug|bottle|สติ๊กเกอร์|น้ำยาล้างเครื่องมือ|sterile glove|ไหม/)) return '7 เวชภัณฑ์';
  if (/[฀-๿]/.test(raw.replace(/\(.*?\)/g, '')) && !/^(unison|ammonium|methyl)/i.test(raw)) return '5 สมุนไพร';
  if (has(/\b(inj|injection|vaccine|vial|amp|penfill|cartridge|ringer|d5s|d5w|d5s\/2)\b|insulin/)) return '1 ยาฉีด';
  if (has(/\b(cream|cre|oint|ointment|gel|gelly|lotion|lot|eye|ear|nasal|spray|mdi|evohaler|inhaler|nebule|inhalation|scrub|irrigate|irrigation|mouth wash|suppo|vag|oral paste|paste|enema|spirit|spi|paint)\b|unison|aromatic ammonia|benzyl|calamine|silver sulfadiazine|hydrogen peroxide|povidine|povidone|ethanol|alcohol|chlorhexidine|hypochlorite|salicylic|tears|soln 0\.5%|soln for inhalation/)) return '2 ยาใช้ภายนอก';
  if (has(/\b(syr|syrup|susp|sus|suspension|mixt|mix|mixture|drop|drops|dro|elixir|lactulose|milk of magnesia|ors|powder|pow)\b/)) return '3 ยาน้ำ';
  if (CHRONIC_.test(k)) return '6 ยาโรคเรื้อรัง';
  if (has(/\b(tab|tablet|cap|capsule|ec|แผง|chewable)\b/)) return '4 ยาเม็ด';
  return '7 เวชภัณฑ์';
}

/* เมนู คลังใน → จัดหมวดยาใหม่ (8 หมวดแบบบ้านเป้า): เขียนช่อง "หมวด/ชั้น" ใหม่ทุกแถว ไม่แตะรหัส ชื่อ หรือยอด */
function applyBanpaoShelves(force) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.items);
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  if (force !== true) {
    try {
      const ui = SpreadsheetApp.getUi();
      const ans = ui.alert('จัดหมวดยาใหม่เป็น 8 หมวดแบบบ้านเป้า', 'ระบบจะเขียนช่อง "หมวด/ชั้น" ใหม่ทุกแถว (' + n + ' รายการ) หมวดที่แก้เองไว้จะถูกแทนที่\nรหัส ชื่อยา QR และยอดคงเหลือไม่เปลี่ยน ต้องการทำต่อไหม', ui.ButtonSet.YES_NO);
      if (ans !== ui.Button.YES) return 0;
    } catch (_) {}
  }
  const v = sh.getRange(2, 1, n, 9).getValues();
  const out = v.map(r => [r[1] || r[7] ? guessShelf_(r[7] || r[1], r[3]) : r[4]]);
  sh.getRange(2, 5, n, 1).setValues(out);
  const cnt = {}; out.forEach(o => { cnt[o[0]] = (cnt[o[0]] || 0) + 1; });
  const msg = 'จัดหมวดใหม่แล้ว ' + n + ' รายการ\n' + SHELVES.map(x => x + ': ' + (cnt[x] || 0)).join('\n') + '\n\nตรวจแล้วแก้เองได้ในช่อง "หมวด/ชั้น" · ถ้ามีแท็บนับยอด ให้กด "สร้างแท็บนับยอด" ใหม่เพื่อเรียงตามหมวด';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
  return n;
}

/* ดึงยาที่ยังไม่มีในแท็บรายการยาจาก DrugMaster และอัปเดตทุน/หน่วยจาก StockLots */
function syncFromHub(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET.items);
  if (!sh) sh = ss.insertSheet(SHEET.items);
  sh.getRange(1, 1, 1, ITEM_HEADER.length).setValues([ITEM_HEADER]).setFontWeight('bold').setBackground('#e0f1ec');
  sh.setFrozenRows(1);
  const hub = hubSS_();
  const M = hubTable_(hub, HUB.master);
  const cN = M.col('name'), cU = M.has('unit') ? M.col('unit') : -1, cA = M.has('is_active') ? M.col('is_active') : -1;
  const cost = {};
  hubLots_(hub).forEach(l => { if (l.cost && (l.qty > 0 || !(l.name in cost))) cost[l.name] = l.cost; });
  const n = sh.getLastRow() - 1;
  const cur = n > 0 ? sh.getRange(2, 1, n, ITEM_HEADER.length).getValues() : [];
  const key = s => String(s || '').trim().toLowerCase();
  const have = {};
  let max = 0;
  cur.forEach(r => {
    if (r[7]) have[key(r[7])] = 1; if (r[1]) have[key(r[1])] = 1;
    const m = new RegExp('^' + CODE_PREFIX + '(\\d+)$').exec(String(r[0]).trim()); if (m) max = Math.max(max, Number(m[1]));
  });
  const add = [];
  M.vals.slice(1).forEach(r => {
    const name = String(r[cN] || '').trim();
    if (!name || have[key(name)]) return;
    have[key(name)] = 1;
    const on = cA < 0 ? true : !(r[cA] === false || String(r[cA]).toUpperCase() === 'FALSE');
    add.push([CODE_PREFIX + String(++max).padStart(CODE_DIGITS, '0'), name, '', cU < 0 ? '' : String(r[cU] || '').trim(), guessShelf_(name, cU < 0 ? '' : r[cU]), '', on, name, cost[name] || '']);
  });
  if (cur.length) {
    const col = cur.map(r => { const c = cost[String(r[7] || r[1] || '').trim()]; return [c || r[8] || '']; });
    sh.getRange(2, 9, col.length, 1).setValues(col);
  }
  if (add.length) {
    const r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, add.length, ITEM_HEADER.length).setValues(add);
    sh.getRange(r0, 7, add.length, 1).insertCheckboxes();
  }
  sh.getRange('I:I').setNumberFormat('#,##0.00');
  const msg = add.length
    ? `เพิ่มยาใหม่จากคลังยาใหญ่ ${add.length} รายการ (${add[0][0]} ถึง ${add[add.length - 1][0]})\n\nพิมพ์สติกเกอร์: เลือกแถวยาใหม่ในแท็บรายการยา แล้วกด คลังใน → พิมพ์สติกเกอร์ QR`
    : 'รายการยาตรงกับคลังยาใหญ่แล้ว ไม่มียาใหม่ (อัปเดตทุน/หน่วยให้แล้ว)';
  Logger.log(msg);
  if (silent !== true) { try { SpreadsheetApp.getUi().alert(msg); } catch (_) {} }
  return add.length;
}

/* ตัวตั้งเวลา: ดึงยาใหม่ทุกคืนตอนตี 1 (กดรันครั้งเดียว) */
function nightlyHubSync() { syncFromHub(true); }
function installHubSync() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'nightlyHubSync').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('nightlyHubSync').timeBased().everyDays(1).atHour(1).inTimezone(TZ).create();
  Logger.log('ตั้งเวลาดึงยาใหม่จากคลังยาใหญ่ทุกคืนแล้ว');
}
