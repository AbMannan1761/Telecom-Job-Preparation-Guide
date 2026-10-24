# টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং গুরুত্বপূর্ণ ফর্মুলা শীট (Essential Formulas Cheatsheet)

চাকরির লিখিত ও প্রিলিমিনারি পরীক্ষায় দ্রুত গাণিতিক সমস্যার সমাধান করার জন্য সবচেয়ে জরুরি সূত্রসমূহ:

---

## ১. সিগন্যাল ও কমিউনিকেশন সিস্টেম (Communication Systems)

### ক. শ্যানন-হার্টলি উপপাদ্য (Shannon Channel Capacity)
$$C = B \log_2 \left(1 + \frac{S}{N}\right) \quad \text{[bps]}$$
- $C$ = চ্যানেল ক্যাপাসিটি (Bits per second)
- $B$ = চ্যানেলের ব্যান্ডউইথ (Hz)
- $S/N$ = সিগন্যাল টু নয়েজ রেশিও (লিনিয়ার মান, dB নয়)।  
  *মনে রাখবেন:* $\text{SNR (dB)} = 10 \log_{10}(S/N) \implies S/N = 10^{\frac{\text{SNR (dB)}}{10}}$

### খ. নাইকুইস্ট বিট রেট (Nyquist Bit Rate for Noiseless Channel)
$$R = 2B \log_2(M) \quad \text{[bps]}$$
- $M$ = ডিসক্রিট সিগন্যাল লেভেল সংখ্যা (যেমন: বাইনারির জন্য $M=2$)

### গ. স্যাম্পলিং থিওরেম ও পিসিএম (PCM & Sampling)
- **নাইকুইস্ট স্যাম্পলিং রেট:** $f_s \ge 2 f_m$
- **কোয়ান্টাইজেশন লেভেল ও বিট সংখ্যা:** $L = 2^n \implies n = \log_2(L)$
- **পিসিএম বিট রেট ($R_b$):** $R_b = n \cdot f_s = n \cdot 2 f_m$
- **ন্যূনতম ট্রান্সমিশন ব্যান্ডউইথ ($BW$):** $BW_{\min} = \frac{R_b}{2} = n \cdot f_m$
- **পিসিএম সিগন্যাল-টু-কোয়ান্টাইজেশন নয়েজ রেশিও ($SQNR$):**
  $$SQNR_{\text{dB}} \approx 6.02 n + 1.76 \text{ dB}$$

---

## ২. সেলুলার ও মোবাইল কমিউনিকেশন (Cellular Network Math)

### ক. ক্লাস্টার সাইজ ও রিইউজ ডিস্ট্যান্স
- **সম্ভাব্য ক্লাস্টার সাইজ ($N$):** 
  $$N = i^2 + ij + j^2 \quad (i, j \text{ অঋণাত্মক পূর্ণসংখ্যা; যেমন: } 1, 3, 4, 7, 9, 12, \dots)$$
- **কো-চ্যানেল রিইউজ রেশিও ($Q$):**
  $$Q = \frac{D}{R} = \sqrt{3N}$$
  - $D$ = নিকটবর্তী কো-চ্যানেল সেলের মধ্যবর্তী দূরত্ব
  - $R$ = সেলের ব্যাসার্ধ (Cell Radius)

### খ. ট্রাফিক পরিমাপ (Traffic Engineering - Erlang)
- **ট্রাফিক তীব্রতা ($A$):**
  $$A = \lambda \cdot H \quad \text{[Erlang]}$$
  - $\lambda$ = প্রতি ঘণ্টায় কল আসার গড় হার (Call arrival rate)
  - $H$ = প্রতিটি কলের গড় স্থায়িত্বকাল (Average holding time in hours)

---

## ৩. অপটিক্যাল ফাইবার কমিউনিকেশন (Optical Fiber Formulas)

### ক. নিউমেরিক্যাল অ্যাপারচার ও এক্সেপটেন্স কোণ
- **নিউমেরিক্যাল অ্যাপারচার ($NA$):**
  $$NA = \sqrt{n_1^2 - n_2^2} \approx n_1 \sqrt{2\Delta}$$
  - $n_1$ = কোরের প্রতিসরাঙ্ক (Core refractive index)
  - $n_2$ = ক্ল্যাডিংয়ের প্রতিসরাঙ্ক (Cladding refractive index)
  - $\Delta = \frac{n_1 - n_2}{n_1}$ = রিলেটিভ রিফ্র্যাক্টিভ ইনডেক্স ডিফারেন্স
- **এক্সেপটেন্স কোণ ($\theta_a$):**
  $$\theta_a = \sin^{-1}(NA)$$

### খ. অপটিক্যাল লিংক পাওয়ার বাজেট
$$P_{rx} = P_{tx} - \text{Total Losses} - \text{Safety Margin}$$
$$\text{Total Losses} = (\alpha \cdot L) + (N_s \cdot \text{Loss}_s) + (N_c \cdot \text{Loss}_c)$$
- $P_{tx}$ = অপটিক্যাল ট্রান্সমিটার পাওয়ার (dBm)
- $\alpha$ = ফাইবার অ্যাটেনুয়েশন সহগ (dB/km)
- $L$ = ফাইবারের মোট দৈর্ঘ্য (km)
- $N_s, \text{Loss}_s$ = স্প্লাইস সংখ্যা ও প্রতি স্প্লাইসে লস (dB)
- $N_c, \text{Loss}_c$ = কানেক্টর সংখ্যা ও প্রতি কানেক্টরে লস (dB)

---

## ৪. আরএফ ও মাইক্রোওয়েভ লিঙ্ক বাজেট (RF & Microwave Formulas)

### ক. ফ্রি স্পেস পাথ লস (Free Space Path Loss - FSPL)
$$FSPL (\text{dB}) = 32.44 + 20 \log_{10}(f_{\text{MHz}}) + 20 \log_{10}(d_{\text{km}})$$
বা, গিগাহার্টজে:
$$FSPL (\text{dB}) = 92.45 + 20 \log_{10}(f_{\text{GHz}}) + 20 \log_{10}(d_{\text{km}})$$

### খ. ফ্রেনেল জোন ব্যাসার্ধ (First Fresnel Zone Radius)
$$r_1 = 17.32 \sqrt{\frac{d_1 \cdot d_2}{f \cdot D}} \quad \text{[meters]}$$
- $d_1, d_2$ = অবস্ট্যাকল থেকে যথাক্রমে ট্রান্সমিটার ও রিসিভারের দূরত্ব (km)
- $D = d_1 + d_2$ = মোট লিঙ্ক দূরত্ব (km)
- $f$ = ফ্রিকোয়েন্সি (GHz)
- *লক্ষ্যণীয়:* ক্লিয়ারেন্স ন্যূনতম ৬০% ($0.6 \times r_1$) নিশ্চিত করতে হয়।

### গ. অ্যান্টেনা প্যারামিটারস
- **EIRP:** $\text{EIRP (dBm)} = P_{tx} (\text{dBm}) - \text{Cable Loss (dB)} + G_{tx} (\text{dBi})$
- **রিসিভড সিগন্যাল লেভেল (RSL):**
  $$RSL (\text{dBm}) = \text{EIRP (dBm)} - FSPL (\text{dB}) + G_{rx} (\text{dBi}) - \text{Rx Cable Loss (dB)}$$
