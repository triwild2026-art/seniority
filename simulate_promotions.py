#!/usr/bin/env python3
"""
simulate_promotions.py

Simulate retirements and promotions per the user's rules and write CSV and XLSX outputs:
- retirements_promotions.csv
- retirements_promotions.xlsx

Run locally: python simulate_promotions.py

This script is intended to be run by the repository workflow as well.
"""
from datetime import datetime, timedelta
import re
import csv
from io import StringIO
import pandas as pd

# Final seniority list provided by the user (columns: Seniority No, Name of Post, Name of Employee, Date of retirement)
# Note: some dates use '.' or '-' as separators; parser normalizes both.
data_text = """1	Sr.AA	 Maya Gopinath	 31-05-2027
2	Sr.AA	 Pravitha K S	 31-05-2027
3	Sr.AA	 Pramod C	 31-05-2027
4	Sr.AA	 SHYMA P	 31-12-2027
5	Sr.AA	 Manoj K M	 31-01-2028
6	Sr.AA	 Beena S	 31-05-2028
7	Sr.AA	 Rajesh C J	 31-05-2028
8	Sr.AA	 Rajesh Kumar R	 31-05-2028
9	Sr.AA	 Preetha O K	 31-08-2028
10	AA	 Salil P S	 31-05-2028
11	AA	 Geethopan G	 31-05-2027
12	AA	 Hobby M V	 31-05-2029
13	AA	 Beena K V	 31-12-2029
14	AA	 Mini Mol Augustin	 30-04-2030
15	AA	 SREEREKHA S	 31-05-2030
16	AA	 Rajee V Kumar	 31-10-2027
17	AA	 Pradeep P P	 30-04-2030
18	AA	 Nandakumar B	 28-02-2031
19	AA	 Suji Stanly	 28-02-2031
20	AA	 Beena P C	 31-05-2030
21	AA	 Neena James	 31-05-2030
22	SS	 Biju S L	 31-05-2030
23	SS	 Chitra P	 30-06-2030
24	SS	 Nazeer S	 31-05-2027
25	SS	 Anil K Sekhar	 31-05-2027
26	SS	 Shijo C Kurian	 30-09-2027
27	SS	 Indira K M	 31-08-2028
28	SS	 Jisha Abraham	 31-05-2033
29	SS	 Leena Nadesh	 31-05-2033
30	SS	 Anil Kumar G	 31-05-2027
31	SS	 Joshy Prasoon N M	 31-05-2032
32	SS	 Harikrishnan S	 31-05-2028
33	SS	 Anilda Thomas	 31-05-2053
34	SS	 Jayakrishnan K G	 30-11-2032
35	SS	 Susmitha M V	 30-04-2033
36	SS	 Sulabha S L	 31-05-2035
37	SS	 Naizamudheen S	 31-05-2030
38	SS	 Manojkumar M	 31-05-2034
39	SS	 Hijas TA	 31-05-2028
40	SS	 Satheesan M K	 31-05-2029
41	SS	 Simy B	 31-05-2028
42	SS	 Jyothi P	 31-05-2032
43	SS	 Mujeeb Rahman PK	 31-10-2027
44	SS	 Prasad NV	 31-05-2035
45	SS	 Minimol KM	 31-05-2027
46	SS	 Vipitha PV	 31-05-2036
47	SS	 Lekha MP	 30-04-2032
48	SS	 Divya K	 28-02-2037
49	SS	 Rajalekshmi KR	 31-05-2029
50	SS	 SujishKumar V R	 31-05-2033
51	SS	 SanthoshKumar K	 31-05-2030
52	SS	 Dileep EN	 28-02-2030
53	SS	 TresaRajan	 30-09-2028
54	SS	 Aji P	 31-05-2032
55	SS	 Suresh Kumar CM	 30-11-2027
56	JS	 Biju K.M	 31-05-2027
57	JS	 Savad M.A	 28-02-2030
58	JS	 Geetha Narayanan.P	 31-01-2034
59	JS	 Bejoy Bhaskaran	 31-01-2034
60	JS	 Manoj Kumar K.P	 30-04-2034
61	JS	 Pradeep Paul	 31-12-2029
62	JS	 Saji P George	 31-05-2028
63	JS	 Rajesh M.V	 31-05-2033
64	JS	 Sreekanth E	 31-05-2031
65	JS	 Biji Sivan	 31-05-2030
66	JS	 Manoj K.N	 30-11-2034
67	JS	 Solly Mol V.M	 31-05-2033
68	JS	 Bindhu E	 31-05-2028
69	JS	 Nisha P.R	 31-03-2037
70	JS	 Ramadas K	 31-05-2027
71	JS	 Manoj S	 31-05-2027
72	JS	 Shaji V	 31-10-2031
73	JS	 Priya K	 30-11-2032
74	JS	 Siad Khan P.U	 31-05-2029
75	JS	 Kunjachan P.R	 31-05-2030
76	JS	 Muhammed Saly K.A	 31-05-2028
77	JS	 Valsaraj A.V	 30-11-2032
78	JS	 Anil Kumar S	 31-05-2030
79	JS	 Santhosh Kumar M	 31-05-2027
80	JS	 Jayasree G	 29-02-2028
81	JS	 Jubily Jerome	 31-10-2033
82	JS	 Moideen V	 31-05-2033
83	JS	 Chandran P	 31-05-2028
84	JS	 Binoy A.M	 30-04-2030
85	JS	 Anvar Sadath A	 31-05-2033
86	JS	 Sajimon O H	 31-05-2036
87	JS	 Sobhana A	 31-01-2028
88	JS	 Abdul Sathar Lebba A	 31-05-2029
89	JS	 Santhamma V	 31-05-2027
90	JS	 Dileep Kumar S	 30-04-2031
91	JS	 Shaju N.O	 31-03-2029
92	JS	 Nisheed S.B	 31-01-2032
93	JS	 Jiju M.G	 30-11-2029
94	JS	 Sudeer.U	 28-02-2030
95	JS	 Jayakumar B	 31-05-2029
96	JS	 Abdul Nazar P.K	 31-12-2031
97	JS	 Smithi S.P	 31-05-2036
98	JS	 Dileep B	 31-12-2030
99	JS	 Aju V R	 31-12-2035
100	JS	 Binoy J P	 31-05-2031
101	JS	 Divakaran N	 30-04-2029
102	JS	 Sivakumar K	 31-05-2028
103	JS	 Hariharan M	 31-05-2031
104	JS	 Radhakrishnan K	 31-05-2027
105	JS	 Muraleedharan P R	 30-11-2028
106	JS	 Vrinda S	 30-04-2038
107	JS	 Manoj Mathew	 29-02-2028
108	JS	 Preethi G Gopi	 31-03-2028
109	JS	 Ashish US	 31-05-2039
110	JS	 Prasanth K K	 31-05-2032
111	JS	 Harilal DS	 31-05-2028
112	JS	 Remya K V	 31-05-2038
113	JS	 Prabha KP	 31-05-2038
114	JS	 Jilu Mary John	 31-08-2036
115	JS	 Ajith Kumar T	 31-05-2027
116	JS	 Sajo B Netto	 31-05-2035
117	JS	 Shanavas K	 31-05-2030
118	JS	 Jayasree R	 31-05-2031
119	JS	 Vinod Kumar K	 31-05-2030
120	JS	 Suja M.O	 30-06-2028
121	JS	 Biju M.K	 30-04-2029
122	JS	 Suresh Babu B	 31-05-2032
123	JS	 Muralikrishnan K C	 31-05-2032
124	JS	 Basheer PM	 31-08-2029
125	JS	 Mini George	 31-08-2028
126	JS	 James Kurian	 31-03-2032
127	JS	 Sanju D S Nair	 31-05-2038
128	HA	 Sindhu M.K	 31.05.2030
129	HA	 Sreejith A P	 31.05.2029
130	HA	 Rajeswaran Nair N	 31.05.2030
131	HA	 Vinod Kumar K.G	 21.05.2030
132	HA	 Anil V Nair	 31.05.2031
133	HA	 Manju S Gopi	 31.10.2027
134	HA	 Premjith D	 31.05.2031
135	HA	 Sankaranarayana Iyer R	 30.09.2030
136	HA	 Sabu P T	 28-02-2031
137	HA	 Biju K R	 29.02-2032
138	HA	 Siji Simon	 30.09.2034
139	HA	 Sandhya S	 31.05.2035
140	HA	 Lekshmi P Nair	 30.09-2040
141	HA	 Aji V	 31.05.2028
142	HA	 Babitha P K	 31.05.2028
143	HA	 Shabina V K	 31.05.2028
144	HA	 Rejitha T	 28.02-2029
145	HA	 Shallin M	 31.05-2028
146	HA	 Sreeraj K K	 31.08-2028
147	HA	 Reghunath A T	 31.12-2028
148	HA	 Manoj G	 31.01-2029
149	HA	 Sajeesh K.P	 31.05.2028
150	HA	 Muhammed Abdul Vahab K.P	 31.05.2036
151	HA	 Anoop Easo	 31.05.2033
152	HA	 Shyam Kumar K	 30.04-2040
153	HA	 Salila V R	 31.05.2029
154	HA	 Jojo Chandran J	 30.04-2035
155	HA	 Anju	 31.05.2035
156	HA	 Kiran C	 31.05.2032
157	HA	 Radha M	 31.05.2028
158	HA	 Sanoj Kumar K S	 30.11-2029
159	HA	 Rajesh G K	 31.05-2028
160	HA	 Sunish V Verol	 30.04.2031
161	HA	 Syam S	 31-05-2039
162	HA	 Abdussammed T M	 31.03.2032
163	HA	 Muhammed Shareef O	 30.04-2034
164	HA	 Nishad P.Muhammed	 31.05.2032
165	HA	 Prajesh P.N	 31.05.2036
166	HA	 Raveendran Kallampoyil	 31.12-2032
167	HA	 Vinod M K	 31.03.2037
168	HA	 Semeer K	 30.04-2034
169	HA	 Sabu M	 31.05.2028
170	HA	 Santhosh Kumar B	 30.04-2033
171	HA	 Pramod N T	 30.11-2033
172	HA	 Rajith Chandran R	 30.04-2032
173	HA	 Aneeshkumar M.K	 31.12-2032
174	HA	 Simi Francy N R	 31.05.2041
175	HA	 Salini T K	 31.10.2041
176	HA	 Suresh P K	 31-05-2035
177	HA	 Anil Kumar K	 31.12-2035
178	HA	 Anaz A K	 31.05-2036
179	HA	 Biju K S	 31.05.2030
180	HA	 Mallika K	 31.05.2029
181	HA	 Anilkumar S	 31.05.2030
182	HA	 Jestin Joseph	 31.05.2029
183	HA	 Preetha T C	 31.05.2031
184	HA	 Marykutty Dominic	 31-05-2029
185	HA	 Anoop RB	 30-11-2035
186	HA	 Mathew John	 28-02-2029
187	HA	 Dhaniya S Remanan	 31-05-2036
188	HA	 Manoj V Chettiar	 31-05-2032
189	HA	 Beena A O	 30-11-2029
190	HA	 Gireesh G	 31-05-2027
191	HA	 Ranjan R	 31-05-2035
192	HA	 Najeeb K S	 30-04-2030
193	HA	 Sakkir Hussain A	 31.10.2030
194	HA	 Kanakam S	 31-03-2035
195	HA	 Sudheendran C	 31-05-2030
196	HA	 Dhanesh Babu P P	 31-05-2034
197	HA	 Rajeev V R	 31-05-2031
198	HA	 Krishnanunni P	 28-02-2034
199	HA	 Kutty Krishnan P	 31-03-2034
200	HA	 Dileep Kumar V K	 31-03-2035
201	HA	 Parvathy J	 31-03-2040
"""


def parse_date(s):
    s = s.strip().replace('.', '-')
    m = re.search(r'(\d{1,2})\D+(\d{1,2})\D+(\d{2,4})', s)
    if not m:
        raise ValueError(f"Unrecognized date: {s!r}")
    d = int(m.group(1)); mo = int(m.group(2)); y = int(m.group(3))
    if y < 100: y += 2000
    return datetime(y, mo, d).date()

# Build people list and infer per-post seniority from global Seniority No ordering
people = []
for line in data_text.strip().splitlines():
    parts = [p.strip() for p in line.split('\t') if p.strip()!='']
    if len(parts) < 4:
        continue
    seniority_no = int(parts[0])
    post = parts[1]
    name = parts[2]
    retirement = parse_date(parts[3])
    people.append({
        "Seniority No": seniority_no,
        "Name": name,
        "Present Post": post,
        "Retirement Date": retirement,
        "Seniority": None,
        "Promotion to JS": None,
        "Promotion to SS": None,
        "Promotion to AA": None,
        "Promotion to Sr.AA": None,
    })

# Infer Seniority-in-post: for each post, order by Seniority No ascending and assign 1,2,3...
from collections import defaultdict
by_post = defaultdict(list)
for p in sorted(people, key=lambda x: x["Seniority No"]):
    by_post[p["Present Post"]].append(p)
for post, lst in by_post.items():
    for idx, p in enumerate(lst, start=1):
        p["Seniority"] = idx

# Promotion mapping
lower_of = {"Sr.AA":"AA", "AA":"SS", "SS":"JS", "JS":"HA"}

# Retirement events sorted by date asc, then Seniority No asc
events = sorted(people, key=lambda p: (p["Retirement Date"], p["Seniority No"]))

def find_eligible_candidate(post, promotion_date):
    candidates = [q for q in people if q.get("Present Post") == post and q["Retirement Date"] > promotion_date]
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x["Seniority"], x["Seniority No"]))
    return candidates[0]

# Process retirements sequentially
for ev in events:
    retire_date = ev["Retirement Date"]
    promotion_date = retire_date + timedelta(days=1)
    vac_post = ev["Present Post"]
    ev["Present Post"] = None  # they retired immediately and are no longer eligible
    if not vac_post:
        continue
    # cascade promotions upwards until HA (stop)
    while vac_post in lower_of:
        from_post = lower_of[vac_post]
        candidate = find_eligible_candidate(from_post, promotion_date)
        if candidate is None:
            break
        # record promotion date
        col = None
        if vac_post == "JS": col = "Promotion to JS"
        elif vac_post == "SS": col = "Promotion to SS"
        elif vac_post == "AA": col = "Promotion to AA"
        elif vac_post == "Sr.AA": col = "Promotion to Sr.AA"
        if col and candidate[col] is None:
            candidate[col] = promotion_date
        # move candidate up
        candidate["Present Post"] = vac_post
        vac_post = from_post
        if vac_post == "HA":
            break

# Prepare output rows
cols = ["Seniority No","Name","Present Post","Seniority","Retirement Date",
        "Promotion to JS","Promotion to SS","Promotion to AA","Promotion to Sr.AA"]

rows = []
for p in sorted(people, key=lambda x: x["Seniority No"]):
    def fmt(d):
        return "" if d is None else d.strftime("%d-%m-%Y")
    rows.append([
        p["Seniority No"],
        p["Name"],
        p.get("Present Post") or "",
        p.get("Seniority"),
        fmt(p["Retirement Date"]),
        fmt(p["Promotion to JS"]),
        fmt(p["Promotion to SS"]),
        fmt(p["Promotion to AA"]),
        fmt(p["Promotion to Sr.AA"]),
    ])

# Create DataFrame and write CSV & XLSX
df = pd.DataFrame(rows, columns=cols)
csv_path = "retirements_promotions.csv"
xlsx_path = "retirements_promotions.xlsx"
df.to_csv(csv_path, index=False)
df.to_excel(xlsx_path, index=False)

print(f"Wrote: {csv_path} and {xlsx_path}")
