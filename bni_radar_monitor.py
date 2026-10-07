import os
import json
import re
import time
import imaplib
import email
import html as html_module
from email.header import decode_header
import requests
from datetime import datetime, timezone
from urllib.parse import quote

# ── CONFIG ─────────────────────────────────────────────────────────────────────
OUTLOOK_EMAIL    = os.environ.get("OUTLOOK_EMAIL", "")
OUTLOOK_PASSWORD = os.environ.get("OUTLOOK_PASSWORD", "")
CALLMEBOT_PHONE  = os.environ.get("CALLMEBOT_PHONE", "447805542098")
CALLMEBOT_APIKEY = os.environ.get("CALLMEBOT_APIKEY", "6038352")
GROUPS_JSON      = os.environ.get("FB_GROUPS", '[{"id":"1996301537277394","name":"Horley Life"}]')

# ── MEMBER + KEYWORD DATABASE ──────────────────────────────────────────────────
MEMBERS = [
  {"name":"Andrew Watkins","company":"The WatkinsHamilton Group","trade":"Air Conditioning","phone":"03302237820","keywords":["air conditioning","air con","AC unit","AC installation","AC repair","AC service","air con not working","split unit","split system","Daikin","Mitsubishi","Samsung","Panasonic","Fujitsu","Toshiba","commercial air conditioning","office air conditioning","server room cooling","too hot in office","office too warm","HVAC","climate control","heat recovery","comfort cooling","air con maintenance","refrigerant","F-gas","ceiling cassette","fan coil","ducted air con","ventilation","MVHR"]},
  {"name":"Anne-Marie Fowler","company":"The Priory Law Group","trade":"Solicitor / Legal","phone":"01737952305","keywords":["solicitor","lawyer","legal advice","conveyancing","house purchase","selling my house","property sale","lease","tenancy agreement","employment dispute","tribunal","contract dispute","legal letter","court","litigation","personal injury","divorce","separation","family law","probate","inheritance","business legal","company law","shareholder agreement","NDA","need a solicitor","looking for legal advice","recommend a lawyer","GDPR","employment contract"]},
  {"name":"Anthony Jacks","company":"Merstham Glass","trade":"Glazier / Glass","phone":"01737644011","keywords":["glazier","glass","broken window","cracked glass","smashed window","window repair","double glazing","triple glazing","window replacement","new windows","shower screen","glass door","glass partition","splashback","mirror","frameless glass","balustrade glass","conservatory glass","emergency glazier","boarded up window","misted glass","condensation between panes"]},
  {"name":"Antony Gray","company":"O&M Office Equipment","trade":"Office Equipment / Printers","phone":"0208 643 4481","keywords":["photocopier","printer","copier","managed print","print lease","office printer","printer lease","toner","cartridge","printer not working","copier repair","office equipment","printer rental","document management","upgrade printer","photocopier contract","multifunction printer","MFP"]},
  {"name":"Ashley Thaw","company":"A1 Heating & Plumbing","trade":"Plumber / Heating Engineer","phone":"0203 928 0941","keywords":["plumber","plumbing","boiler","boiler service","boiler repair","boiler breakdown","no hot water","no heating","central heating","radiator","radiator leak","leaking pipe","burst pipe","dripping tap","blocked toilet","toilet repair","bathroom fitting","bathroom renovation","shower installation","underfloor heating","gas safe","gas engineer","combi boiler","system boiler","water pressure","hot water cylinder","immersion heater","pipe repair","emergency plumber"]},
  {"name":"Connor Seagroatt","company":"AC Electrical Surrey","trade":"Electrician","phone":"07948 793713","keywords":["electrician","electrical","rewire","full rewire","consumer unit","fuse box","fuse board","fuseboard upgrade","RCD","circuit breaker tripping","power cut","no power","socket not working","light not working","electrical fault","electrical inspection","EICR","EV charger","electric car charger","home charger","pod point","outdoor lighting","garden lights","security lights","flood lights","additional sockets","new sockets","storage heaters","electric heating","electric shower","PAT testing","electrical certificate","commercial electrical","distribution board"]},
  {"name":"Dan Hussey","company":"DH Tree Services","trade":"Tree Surgeon / Arborist","phone":"07872615504","keywords":["tree surgeon","arborist","tree removal","fell a tree","cut down tree","tree pruning","crown reduction","tree trimming","lopping","pollarding","stump grinding","stump removal","tree stump","hedge cutting","hedge trimming","overgrown hedge","leylandii","fallen tree","emergency tree","tree on roof","tree damage","storm damage","dead tree","diseased tree","tree survey","TPO","tree preservation order","ivy removal","wood chipping","firewood"]},
  {"name":"Daniel Atubo","company":"Omatic Group","trade":"Sports Massage Therapist","phone":"07917 165901","keywords":["sports massage","massage therapist","massage","deep tissue massage","sports injury","muscle injury","muscle pain","back pain","neck pain","shoulder pain","tight muscles","muscle tension","knots","trigger point","injury recovery","injury rehab","rehabilitation","running injury","gym injury","football injury","sports recovery","recovery massage","stress relief","relaxation massage","sciatica","hamstring","calf strain","recommend a massage","sports therapist"]},
  {"name":"David Rutter","company":"TBA Finance","trade":"Finance / Business Loans","phone":"+447565426138","keywords":["business loan","finance","funding","invoice finance","asset finance","bridging loan","bridging finance","commercial mortgage","development finance","cash flow","working capital","business funding","startup funding","equipment finance","vehicle finance","fleet finance","invoice factoring","factoring","lending","need finance","looking for funding","raise finance","capital"]},
  {"name":"Dean Slade","company":"Book The World","trade":"Travel Agent","phone":"01737 952444","keywords":["travel agent","holiday","book a holiday","flights","hotel","corporate travel","business travel","group travel","team trip","incentive travel","cruise","villa","all inclusive","honeymoon","wedding trip","anniversary trip","family holiday","ski trip","skiing","safari","long haul","short break","travel insurance","luxury travel","tailor made holiday"]},
  {"name":"Dean Ward","company":"WPS UK Ltd","trade":"Property Maintenance","phone":"07730982359","keywords":["property maintenance","handyman","repairs","general repairs","bathroom renovation","kitchen renovation","refurbishment","redecoration","tiling","tile","grouting","floor tiles","wall tiles","painting and decorating","decorator","painting","wallpaper","door fitting","window fitting","flat maintenance","landlord","rental property","HMO","office maintenance","building maintenance"]},
  {"name":"Duane W-King","company":"Concept Claims Solutions","trade":"Insurance Loss Assessor","phone":"07944097502","keywords":["insurance claim","loss assessor","loss adjuster","insurance dispute","flood damage","fire damage","storm damage","subsidence claim","escape of water","burst pipe insurance","water damage claim","insurance payout","claim rejected","underpaid claim","insurer dispute","home insurance claim","commercial insurance claim","building insurance","business interruption","theft claim","help with insurance","insurance not paying","claim denied"]},
  {"name":"Harry Wood","company":"Kent Fencing and Landscaping","trade":"Fencing & Landscaping","phone":"07375542927","keywords":["fencing","fence","new fence","fence repair","fence panel","fence post","garden gate","timber gate","wooden fence","close board fence","landscaping","garden design","garden makeover","landscape gardener","patio","patio installation","paving","block paving","driveway","new driveway","resin driveway","tarmac driveway","lawn","turf","new lawn","lawn care","artificial grass","astroturf","raised beds","retaining wall","garden wall","decking","composite decking"]},
  {"name":"Holly Woolford","company":"Ebbisham Financial Solutions","trade":"Financial Advisor / Mortgage Broker","phone":"07816435735","keywords":["mortgage","mortgage broker","mortgage advisor","first time buyer","remortgage","mortgage renewal","buy to let","BTL mortgage","mortgage rate","fixed rate","tracker mortgage","life insurance","critical illness","income protection","financial advisor","financial advice","pension","retirement planning","ISA","investment","savings","wealth management","mortgage declined","bad credit mortgage","self employed mortgage","new build mortgage","shared ownership","help to buy"]},
  {"name":"Jack Caffyn","company":"Logody Creative Studios","trade":"Branding / Logo / Graphic Design","phone":"07455010222","keywords":["logo","logo design","logo designer","new logo","rebrand","rebranding","brand identity","branding","graphic design","designer","business cards","flyer","leaflet","brochure","poster","banner","marketing materials","signage","van wrap","vehicle livery","social media graphics","brand colours","colour palette","typography","startup branding","new business identity","visual identity"]},
  {"name":"James Buckett","company":"Clear It","trade":"House / Office Clearance","phone":"07851006148","keywords":["house clearance","clearance","rubbish removal","junk removal","office clearance","skip hire","waste removal","bulky waste","furniture disposal","old furniture","garden clearance","garden waste","probate clearance","bereavement clearance","hoarder clearance","loft clearance","garage clearance","shed clearance","end of tenancy clearance","rubbish collection","white goods disposal","fridge disposal","mattress disposal"]},
  {"name":"Jamie Harkness","company":"Professional Roof Cleaning","trade":"Roof Cleaning / Exterior Cleaning","phone":"07341338645","keywords":["roof cleaning","moss on roof","green roof","algae roof","dirty roof","pressure washing","jet wash","power wash","driveway cleaning","patio cleaning","gutter cleaning","blocked gutters","softwash","biocide treatment","exterior cleaning","render cleaning","conservatory cleaning","solar panel cleaning","moss removal","lichen removal","black streaks","roof treatment"]},
  {"name":"Jo Suleyman","company":"TRUST Building Services","trade":"Building Services / Construction","phone":"01883772772","keywords":["builder","building work","construction","refurbishment","renovation","extension","house extension","rear extension","loft conversion","building regulations","planning permission","plastering","render","dry lining","stud wall","partition wall","groundworks","foundations","fit out","office fit out","commercial refurbishment"]},
  {"name":"Joanne Hall","company":"Hallways","trade":"Estate Agent / Property","phone":"02035363959","keywords":["estate agent","sell my house","selling my house","property for sale","buy a house","house for sale","property valuation","house valuation","lettings","rental property","landlord","tenant","letting agent","property management","find a tenant","rent my house","rent out property","commercial property","office to let","property investment","buy to let","house prices","property market","sold stc"]},
  {"name":"Joe Brooks","company":"Brooks and Burge Films","trade":"Videography / Film Production","phone":"07477578856","keywords":["videographer","video production","corporate video","promo video","promotional video","brand video","company video","testimonial video","social media video","YouTube video","Instagram video","reel","wedding video","event video","drone footage","aerial video","product video","explainer video","filming","video shoot","video editor","video marketing","video content creation"]},
  {"name":"Joe Warnes","company":"Roma-Scar Flooring Studio","trade":"Flooring","phone":"07377537163","keywords":["flooring","floor","carpet","new carpet","carpet fitting","carpet installation","laminate","laminate flooring","engineered wood","hardwood floor","wood flooring","LVT","luxury vinyl","vinyl flooring","karndean","amtico","floor tiles","porcelain tiles","stone floor","commercial flooring","office flooring","floor replacement","floor sanding","floor restoration","parquet","herringbone floor"]},
  {"name":"Jon Skinner","company":"Glow Homes","trade":"Home Energy / Solar / Battery","phone":"0333 112 7272","keywords":["solar panels","solar","photovoltaic","PV","solar installation","battery storage","home battery","Tesla Powerwall","energy storage","EV charger","electric car charger","heat pump","air source heat pump","ASHP","energy efficiency","energy saving","green energy","renewable energy","electricity bills","reduce energy bills","energy costs","smart meter","energy monitoring","solar grant","ECO4","BUS scheme"]},
  {"name":"Keiran Wynyard","company":"Dreamraven Designs","trade":"Web Design / Digital Design","phone":"07985 410 643","keywords":["website","website design","web designer","new website","website build","WordPress","Shopify","ecommerce","online shop","web development","website redesign","website refresh","landing page","UX design","UI design","app design","mobile app","domain name","hosting","website maintenance","SEO","need a website","website not working","want to sell online"]},
  {"name":"Lee Wilcox","company":"H&H Maintenance – Grounds","trade":"Grounds Maintenance","phone":"07557687007","keywords":["grounds maintenance","lawn mowing","grass cutting","lawn care","hedge trimming","hedge cutting","overgrown garden","garden maintenance","commercial grounds","estate maintenance","grounds contract","weed control","weeding","moss treatment","scarifying","leaf clearing","seasonal maintenance","shrub pruning","border maintenance"]},
  {"name":"Lisa Delaney","company":"Excalibur Bookkeeping","trade":"Bookkeeper","phone":"07850142177","keywords":["bookkeeper","bookkeeping","accounts","VAT return","VAT registration","payroll","payroll processing","PAYE","CIS","construction industry scheme","QuickBooks","Xero","Sage","accounting software","cloud accounting","self assessment","tax return","year end accounts","profit and loss","balance sheet","management accounts","reconciliation","bank reconciliation","invoicing","expense tracking","need a bookkeeper","behind on accounts"]},
  {"name":"Louise Long","company":"The Dylan Strong Foundation","trade":"Charity / Foundation","phone":"07824826998","keywords":["charity","fundraising","donation","sponsor","sponsorship","mental health","young people","foundation","charity event","fundraising event","raffle","auction"]},
  {"name":"Marc Barker","company":"M.E.B Vehicle Services","trade":"Vehicle Servicing / Garage","phone":"02036 334 567","keywords":["car service","vehicle service","MOT","MOT test","car repair","mechanic","tyres","tyre change","puncture","wheel alignment","brake pads","oil change","oil service","engine fault","warning light","engine management","diagnostics","vehicle diagnostics","fleet maintenance","company cars","fleet service","van service","commercial vehicle","bodywork","dent repair","smart repair","pre MOT check","vehicle inspection"]},
  {"name":"Maria Millward","company":"HRoverload","trade":"HR Consultancy","phone":"02085889494","keywords":["HR","human resources","HR advice","employment law","employment contract","disciplinary","grievance","dismissal","redundancy","settlement agreement","TUPE","employee handbook","HR policies","staff handbook","recruitment","hiring process","job description","performance management","appraisal","absence management","maternity","paternity","parental leave","flexible working","HR compliance","employment tribunal"]},
  {"name":"Martin Bullock","company":"Cladmaster","trade":"Cladding / External Wall Systems","phone":"08001389072","keywords":["cladding","external cladding","wall cladding","rainscreen cladding","render","external render","K rend","monocouche","rendering","external insulation","EWI","solid wall insulation","cavity wall insulation","external wall","facade","building envelope","timber cladding","composite cladding","metal cladding","fibre cement","commercial exterior","fire rated cladding","aluminium cladding","panel system"]},
  {"name":"Max Dandy","company":"MD Carpentry","trade":"Carpenter / Joiner","phone":"07825368634","keywords":["carpenter","carpentry","joiner","joinery","fitted furniture","bespoke furniture","fitted wardrobes","wardrobe","built-in wardrobes","alcove units","shelving","fitted kitchen","kitchen installation","kitchen fitting","door fitting","new doors","internal doors","external doors","door frame","skirting boards","architrave","window sill","decking","garden decking","staircase","handrail","balustrade","second fix","studwork","timber frame"]},
  {"name":"Michael Hora","company":"Active Group Scaffolding","trade":"Scaffolding","phone":"07734438937","keywords":["scaffolding","scaffold","scaffolder","scaffold hire","scaffold erection","roof scaffolding","chimney scaffolding","scaffold for painting","scaffold for extension","scaffold for loft conversion","commercial scaffold","industrial scaffold","access scaffold","tower scaffold","mobile scaffold","temporary roof","scaffold inspection"]},
  {"name":"Michael Walker","company":"RLR Leadwork and Roofing","trade":"Roofer / Leadwork","phone":"07821871324","keywords":["roofer","roofing","roof repair","leaking roof","roof leak","roof replacement","re-roofing","new roof","roof tiles","slates","slate roof","flat roof","felt roof","GRP fibreglass roof","EPDM roof","rubber roof","leadwork","lead flashing","flashing repair","valley lead","chimney lead","guttering","gutters","gutter repair","gutter replacement","fascia","soffit","chimney repair","chimney repoint","chimney rebuild","chimney pot","dormer","velux","skylight","storm damage roof","roof survey"]},
  {"name":"Morgan Bowen","company":"Perry Monroe Wealth Management","trade":"Wealth Management / Financial Planning","phone":"01737222646","keywords":["wealth management","financial planning","financial advisor","IFA","pension","retirement","pension review","SIPP","drawdown","annuity","investment","investments","ISA","stocks and shares ISA","portfolio","inheritance tax","IHT","estate planning","trust","power of attorney","life insurance","critical illness","income protection","business protection","key person insurance","shareholder protection","sell my business","exit planning","business succession"]},
  {"name":"Nic Newbury","company":"Newmax Fire & Security","trade":"Fire & Security","phone":"07805542098","keywords":["CCTV","security camera","cameras","fire alarm","fire detection","smoke detector","access control","door entry","key fob","fob access","biometric","facial recognition","intruder alarm","burglar alarm","alarm system","security system","monitored alarm","intercom","video intercom","video doorbell","door phone","entry system","automated gate","electric gate","barrier","gate motor","gate automation","emergency lighting","exit sign","fire extinguisher","fire safety","disabled refuge alarm","smoke ventilation","AOV","fire panel","WiFi","networking","structured cabling","data network","security assessment","security survey","CCTV quote","alarm quote","CCTV not working","alarm fault","Hikvision","Paxton","Dahua","Ajax","Texecom","Honeywell","commercial security","warehouse CCTV","office security","school security"]},
  {"name":"Phil Curtis","company":"Happy Beautiful Health","trade":"Health & Wellbeing","phone":"07921 316963","keywords":["health coach","nutrition","nutritionist","wellbeing","wellness","weight loss","diet","healthy eating","lifestyle change","gut health","workplace wellbeing","employee wellbeing","stress","burnout","mental health","energy levels","fatigue","sleep","hormones","menopause","health programme","wellbeing programme","corporate wellness"]},
  {"name":"Ray Turner","company":"R & M Turner & Co.","trade":"Builder / General Contractor","phone":"0771 281 0229","keywords":["builder","building work","construction","general contractor","extension","house extension","rear extension","side extension","wrap around extension","loft conversion","dormer loft","hip to gable","knock through","open plan","structural wall","RSJ","steel beam","lintel","renovation","refurbishment","house renovation","groundworks","foundation","footings","concrete slab","bricklayer","brickwork","blockwork","masonry","pointing","repointing","garage conversion","outbuilding","garden room","planning permission","building regulations","structural engineer","new build","self build","demolition","porch","utility room","orangery"]},
  {"name":"Ryan Penedo","company":"C4 Plus Drainage","trade":"Drainage","phone":"02036573189","keywords":["drainage","blocked drain","drain blocked","blocked toilet","blocked sink","blocked shower","sewage","sewer","drain survey","CCTV drain survey","drain camera","drain inspection","drain repair","drain excavation","manhole","soakaway","septic tank","cesspit","pump station","root ingress","collapsed drain","drain lining","surface water","flooding","standing water"]},
  {"name":"Sam Coldicott","company":"Coldicott Constructions","trade":"Builder / General Contractor","phone":"07858095919","keywords":["builder","building work","construction","contractor","general contractor","extension","house extension","rear extension","side extension","loft conversion","dormer loft","new build","knock through","open plan","structural wall","RSJ","steel beam","renovation","refurbishment","house renovation","groundworks","foundation","concrete","bricklayer","brickwork","blockwork","masonry","garage conversion","outbuilding","garden room","planning permission","building regulations"]},
  {"name":"Sean Dowd","company":"Optimise Accountants","trade":"Accountant / Tax Advisor","phone":"01293265280","keywords":["accountant","accounting","tax advisor","tax return","self assessment","corporation tax","company accounts","annual accounts","year end","VAT","VAT return","VAT registration","MTD","making tax digital","HMRC","HMRC investigation","tax investigation","payroll","PAYE","CIS","sole trader","limited company","partnership","LLP","R&D tax credits","capital allowances","business tax planning","cashflow","management accounts","profit and loss","Companies House","director salary"]},
  {"name":"Simon Cripps","company":"Smart Cow Marketing","trade":"Marketing / Social Media / Digital","phone":"0203 1371826","keywords":["marketing","social media","digital marketing","marketing agency","Facebook ads","Instagram ads","Google ads","PPC","paid ads","SEO","search engine optimisation","Google ranking","organic search","content marketing","blog","content creation","copywriting","email marketing","newsletter","Mailchimp","social media management","post scheduling","marketing strategy","brand awareness","lead generation","B2B marketing","local SEO","Google My Business","reviews"]},
  {"name":"Simon Roberts","company":"WPA Cedars Health","trade":"Private Health Insurance","phone":"07885 868 260","keywords":["private health insurance","private medical","PMI","health insurance","BUPA","AXA Health","WPA","Vitality","corporate health","employee health benefit","group health insurance","dental insurance","dental plan","cash plan","waiting list","NHS waiting list","private hospital","private treatment","health cash plan","company health plan","team benefits","staff benefits","employee benefits"]},
  {"name":"Steve Drewett","company":"MC Will Writing Services","trade":"Will Writing / Estate Planning","phone":"07386892814","keywords":["will","write a will","will writing","make a will","update my will","power of attorney","lasting power of attorney","LPA","probate","estate planning","inheritance","inheritance tax","trust","family trust","guardianship","executor","beneficiary","bereavement","business will","business LPA","succession planning","living will"]},
  {"name":"Steve Jebson","company":"Business Doctors","trade":"Business Consultancy / Coaching","phone":"07889 269573","keywords":["business consultant","business coach","business advisor","consultancy","business strategy","growth strategy","business planning","business plan","profit improvement","business performance","exit strategy","sell my business","business valuation","MBO","acquisition","management buy out","business sale","due diligence","business mentor","mentoring","coaching","accountability","scaling up","growth","franchise","systems","processes","business challenges","struggling business","turnaround"]},
]

def match_post(post_text):
    text_lower = post_text.lower()
    matches = []
    seen = set()
    for member in MEMBERS:
        for kw in member["keywords"]:
            if kw.lower() in text_lower and member["name"] not in seen:
                matches.append({"member": member, "keyword": kw})
                seen.add(member["name"])
                break
    return matches

def send_whatsapp_alert(group_name, post_text, matches, poster_name="", post_link=""):
    group_url = post_link if post_link else f"https://www.facebook.com/groups/1996301537277394"
    lines = [f"BNI REFERRAL OPPORTUNITY"]
    lines.append(f"Group: {group_name}")
    if poster_name:
        lines.append(f"Posted by: {poster_name}")
    lines.append("")
    for m in matches:
        lines.append(f"Refer: {m['member']['name']} ({m['member']['trade']})")
        lines.append(f"Tel: {m['member']['phone']}")
    lines.append("")
    lines.append(group_url)
    message = "\n".join(lines)
    encoded = quote(message)
    url = f"https://api.callmebot.com/whatsapp.php?phone={CALLMEBOT_PHONE}&text={encoded}&apikey={CALLMEBOT_APIKEY}"
    try:
        resp = requests.get(url, timeout=15)
        print(f"  WhatsApp sent: {resp.status_code}")
    except Exception as e:
        print(f"  WhatsApp error: {e}")

def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                break
            elif ct == "text/html" and not body:
                raw = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                body = re.sub(r'<[^>]+>', ' ', raw)
                body = html_module.unescape(body)
    else:
        raw = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        if '<' in raw:
            body = re.sub(r'<[^>]+>', ' ', raw)
            body = html_module.unescape(body)
        else:
            body = raw
    return ' '.join(body.split())

def check_email_for_posts(email_addr, password, groups):
    posts = []
    group_names = [g.get("name","") for g in groups]

    try:
        print(f"  Connecting to IMAP...")
        mail = imaplib.IMAP4_SSL("mail.newmaxfs.co.uk", 993)
        mail.login(email_addr, password)
        mail.select("INBOX")

        # Search for unread emails from Facebook
        status, messages = mail.search(None, 'UNSEEN FROM "facebookmail.com"')
        if status != "OK" or not messages[0]:
            print("  No new Facebook notification emails")
            mail.logout()
            return []

        ids = messages[0].split()
        print(f"  Found {len(ids)} unread Facebook email(s)")

        for eid in ids:
            status, data = mail.fetch(eid, "(RFC822)")
            if status != "OK":
                continue
            msg = email.message_from_bytes(data[0][1])

            # Decode subject
            raw_subject = msg.get("Subject","")
            subject_parts = decode_header(raw_subject)
            subject = ""
            for part, enc in subject_parts:
                if isinstance(part, bytes):
                    subject += part.decode(enc or 'utf-8', errors='ignore')
                else:
                    subject += part

            # Match group name
            matched_group = "Facebook Group"
            for gname in group_names:
                if gname.lower() in subject.lower():
                    matched_group = gname
                    break

            body = get_email_body(msg)

            # Extract poster name from subject (e.g. "John Smith posted in Horley Life")
            poster = ""
            poster_match = re.match(r'^(.+?)\s+posted in', subject, re.IGNORECASE)
            if poster_match:
                poster = poster_match.group(1).strip()

            # Extract post link from email body
            post_link = ""
            link_match = re.search(r'https://www\.facebook\.com/groups/[^\s"\'<>]+', body)
            if link_match:
                post_link = link_match.group(0)

            if body and len(body) > 20:
                posts.append({"text": body, "subject": subject, "group": matched_group, "poster": poster, "link": post_link})
                print(f"  Email: {subject[:80]}")

            # Mark as read so we don't process it again
            mail.store(eid, '+FLAGS', '\\Seen')

        mail.logout()

    except imaplib.IMAP4.error as e:
        print(f"  Login failed: {e} — check OUTLOOK_EMAIL and OUTLOOK_PASSWORD secrets")
    except Exception as e:
        print(f"  Email error: {e}")

    return posts

def main():
    print(f"BNI Referral Radar starting — {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")

    if not OUTLOOK_EMAIL or not OUTLOOK_PASSWORD:
        print("ERROR: Set OUTLOOK_EMAIL and OUTLOOK_PASSWORD secrets in GitHub.")
        return

    try:
        groups = json.loads(GROUPS_JSON)
    except:
        groups = [{"id":"1996301537277394","name":"Horley Life"}]

    print(f"Checking inbox: {OUTLOOK_EMAIL}")
    posts = check_email_for_posts(OUTLOOK_EMAIL, OUTLOOK_PASSWORD, groups)
    print(f"Posts to process: {len(posts)}")

    total_alerts = 0
    for post in posts:
        matches = match_post(post["text"])
        if matches:
            print(f"  MATCH in '{post['group']}': {[m['member']['name'] for m in matches]}")
            send_whatsapp_alert(post["group"], post["text"], matches, post.get("poster",""), post.get("link",""))
            total_alerts += 1
            time.sleep(3)
        else:
            print(f"  No keyword match in: {post['subject'][:60]}")

    print(f"Done. {total_alerts} alert(s) sent.")

if __name__ == "__main__":
    main()
