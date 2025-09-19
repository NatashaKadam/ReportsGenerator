import os
import hashlib

# --- File Paths and Constants ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH_MERGED = os.path.join(SCRIPT_DIR, "assets", "template_merged.docx")
SSR_DATA_EXCEL = os.path.join(SCRIPT_DIR, "assets", "ssr_data.xlsx")
SESSION_TIMEOUT = 30 * 60 * 1000

# --- User Authentication ---
ADMIN_USER = "admin"
ADMIN_PASS_HASH = hashlib.sha256("bill123".encode()).hexdigest()

# --- Translation Dictionaries ---
TRANSLATIONS = {
    'mr': {
        "Login to Reports Generator": "रिपोर्ट्स जनरेटरमध्ये लॉग इन करा",
        "Username": "वापरकर्तानाव",
        "Password": "पासवर्ड",
        "Login": "लॉग इन करा",
        "Invalid credentials. %d attempts remaining.": "अवैध क्रेडेन्शियल्स. %d प्रयत्न शिल्लक.",
        "Too many failed attempts. Try again in %d seconds.": "खूप जास्त अयशस्वी प्रयत्न. %d सेकंदात पुन्हा प्रयत्न करा.",
        "Account Locked": "अकाउंट लॉक झाले आहे",
        "Session Expired": "सत्र कालबाह्य झाले आहे",
        "Your session has expired. Please login again.": "तुमचे सत्र कालबाह्य झाले आहे. कृपया पुन्हा लॉग इन करा.",
        "Session History": "सत्र इतिहास",
        "New Bill": "नवीन बिल",
        "Search sessions...": "सत्र शोधा...",
        "Document Details": "दस्तऐवज तपशील",
        "Construction Items": "बांधकाम वस्तू",
        "Excess/Saving Statement": "जादा/बचत विवरण",
        "Name": "नाव",
        "Work": "काम",
        "Division": "विभाग",
        "Constituency": "मतदारसंघ",
        "Fund Head": "निधी प्रमुख",
        "Contractor": "कंत्राटदार",
        "Deputy Engineer": "उप अभियंता",
        "Date": "दिनांक",
        "Start Date": "प्रारंभ दिनांक",
        "End Date": "समाप्ती दिनांक",
        "Agreement No": "करार क्र",
        "Work Order No": "कार्य आदेश क्र",
        "Acceptance No": "स्वीकृती क्र",
        "MB No": "एम.बी. क्र",
        "Letter No": "पत्र क्र",
        "Vide Letter No": "व्हिडिओ पत्र क्र",
        "Year": "वर्ष",
        "Est Cost": "अंदाजित खर्च",
        "Amt Rupes": "रुपये रक्कम",
        "Percentage Quoted": "उद्धृत टक्केवारी",
        "Send To": "यांना पाठवा",
        "Subject": "विषय",
        "Message": "संदेश",
        "Edit Message...": "संदेश संपादित करा...",
        "Item Description:": "वस्तूचे वर्णन:",
        "Quantity:": "प्रमाण:",
        "Unit:": "एकक:",
        "Rate:": "दर:",
        "Total Cost:": "एकूण खर्च:",
        "Add Item": "वस्तू जोडा",
        "Jr./Sect./Asst. Engineer:": "कनिष्ठ/सेक्टर/सहाय्यक अभियंता:",
        "Delete": "हटवा",
        "Sr. No": "अनु. क्र",
        "Chapter": "अध्याय",
        "SSR Item No.": "SSR वस्तू क्र.",
        "Reference No.": "संदर्भ क्र.",
        "Description": "वर्णन",
        "Add. Spec.": "अतिरिक्त तपशील.",
        "Unit": "एकक",
        "Rate": "दर",
        "Qty": "प्रमाण",
        "Total": "एकूण",
        "Actions": "क्रिया",
        "Tender Qty": "निविदा प्रमाण",
        "Executed Qty": "अंमलबजावणी केलेले प्रमाण",
        "Excess": "जादा",
        "Saving": "बचत",
        "Remarks": "टीका",
        "Save DOCX": "DOCX जतन करा",
        "Save PDF": "PDF जतन करा",
        "Refresh Preview": "पूर्वावलोकन रीफ्रेश करा",
        "Quick Save": "त्वरित जतन करा",
        "Export Excel": "एक्सेल निर्यात करा",
        "Detach Preview": "पूर्वावलोकन वेगळे करा",
        "Dock Preview": "पूर्वावलोकन डॉक करा",
        "New document ready": "नवीन दस्तऐवज तयार",
        "Unsaved Changes": "न जतन केलेले बदल",
        "You have unsaved changes. Do you want to save them before switching?": "तुमच्याकडे न जतन केलेले बदल आहेत. स्विच करण्यापूर्वी तुम्हाला ते जतन करायचे आहेत का?",
        "Save": "जतन करा",
        "Discard": "सोडून द्या",
        "Cancel": "रद्द करा",
        "Session saved: %s": "सत्र जतन केले: %s",
        "Save failed - check logs": "जतन करणे अयशस्वी - लॉग तपासा",
        "Operation failed: %s": "क्रिया अयशस्वी: %s",
        "File saved to:\n%s": "फाइल येथे जतन केली:\n%s",
        "File saved: %s": "फाइल जतन केली: %s",
        "Preview updated.": "पूर्वावलोकन अद्यतनित केले.",
        "Missing Info": "माहिती गहाळ",
        "Please provide a 'Name' in the Document Details before generating a file.": "कृपया फाइल तयार करण्यापूर्वी 'दस्तऐवज तपशील' मध्ये 'नाव' प्रदान करा.",
        "Export Successful": "निर्यात यशस्वी",
        "Data exported to:\n%s": "डेटा येथे निर्यात केला:\n%s",
        "Exported to Excel: %s": "एक्सेलमध्ये निर्यात केले: %s",
        "Export Failed": "निर्यात अयशस्वी",
        "An error occurred: %s": "एक त्रुटी आली: %s",
        "Error": "त्रुटी",
        "Load Error": "लोड त्रुटी",
        "Could not load session data: %s": "सत्र डेटा लोड करू शकलो नाही: %s",
        "Warning: Attempted to load data from a deleted session item. Ignoring.": "चेतावणी: हटवलेल्या सत्र वस्तूमधून डेटा लोड करण्याचा प्रयत्न केला. दुर्लक्ष करत आहे.",
        "Session list refreshed": "सत्र सूची रीफ्रेश केली",
        "Error refreshing sessions": "सत्र रीफ्रेश करताना त्रुटी",
        "Edit Message": "संदेश संपादित करा",
        "Item Not Found": "वस्तू सापडली नाही",
        "Selected item not in data source.": "निवडलेली वस्तू डेटा स्त्रोतामध्ये नाही.",
        "Invalid Input": "अवैध इनपुट",
        "Enter a positive quantity.": "सकारात्मक प्रमाण प्रविष्ट करा.",
        "Invalid Quantity": "अवैध प्रमाण",
        "Quantity must be a number.": "प्रमाण एक संख्या असावे.",
        "Invalid Item": "अवैध वस्तू",
        "Please select a valid item.": "कृपया एक वैध वस्तू निवडा.",
        "Data Not Loaded": "डेटा लोड झाला नाही",
        "SSR data not available.": "SSR डेटा उपलब्ध नाही.",
        "Excel Data Missing": "एक्सेल डेटा गहाळ",
        "Error: '%s' not found.\nPlease ensure the file exists in the 'assets' folder.": "त्रुटी: '%s' सापडले नाही.\nकृपया 'assets' फोल्डरमध्ये फाइल अस्तित्वात असल्याची खात्री करा.",
        "Invalid Excel File": "अवैध एक्सेल फाइल",
        "Excel file must contain required columns.": "एक्सेल फाइलमध्ये आवश्यक स्तंभ असणे आवश्यक आहे.",
        "Excel Load Error": "एक्सेल लोड त्रुटी",
        "An error occurred while reading the Excel file: %s": "एक्सेल फाइल वाचताना एक त्रुटी आली: %s",
        "Ready": "तयार",
        "Loaded session: %s": "सत्र लोड केले: %s",
        "Unexpected error: %s": "अनपेक्षित त्रुटी: %s",
        "Settings": "सेटिंग्ज",
        "Dark Mode": "डार्क मोड",
        "Software Language": "सॉफ्टवेअर भाषा",
        "English": "इंग्रजी",
        "Marathi": "मराठी",
        "Auto-Save Interval (minutes)": "ऑटो-सेव्ह मध्यांतर (मिनिटे)",
        "Backup & Export Location": "बॅकअप आणि निर्यात स्थान",
        "Choose Location": "स्थान निवडा",
        "No construction items to export.": "निर्यात करण्यासाठी कोणतीही बांधकाम वस्तू नाही.",
        "Report pack generated.": "रिपोर्ट पॅक तयार झाला.",
        "Report pack failed: %s": "रिपोर्ट पॅक अयशस्वी: %s",
        "Report pack generation canceled.": "रिपोर्ट पॅक निर्मिती रद्द केली.",
        "File saved successfully in:\n%s": "फाइल येथे यशस्वीरित्या जतन केली:\n%s",
        "Generating Report Pack...": "रिपोर्ट पॅक तयार करत आहे...",
        "Generating DOCX file...": "DOCX फाइल तयार करत आहे...",
        "Generating PDF file...": "PDF फाइल तयार करत आहे...",
        "Generating Excel file...": "एक्सेल फाइल तयार करत आहे...",
        "Success": "यश",
        "PDF Not Available": "PDF उपलब्ध नाही",
        "`pypandoc` library not installed.": "`pypandoc` लायब्ररी स्थापित नाही.",
        "Dependency Missing": "अवलंबन गहाळ",
        "Cannot generate PDF because `pypandoc` is not installed.": "`pypandoc` स्थापित नसल्यामुळे PDF तयार करू शकत नाही.",
        "DOCX generation failed: %s": "DOCX निर्मिती अयशस्वी: %s"
    }
}

# --- Data for Material Consumption ---
MATERIAL_CONSUMPTION_MAP = {
    "soling": {
        "short_desc": "Soling",
        "ratios": {"sand": 0.000, "rubble": 1.200, "brick": 0.00, "metal": 0.000, "cement": 0.000}
    },
    "s.w. pipe": {
        "short_desc": "9\" GSW Pipe",
        "ratios": {"sand": 0.000, "rubble": 0.000, "brick": 0.00, "metal": 0.000, "cement": 0.080}
    },
    "m15": {
        "short_desc": "P.C.C. 1:2:4",
        "ratios": {"sand": 0.445, "rubble": 0.000, "brick": 0.00, "metal": 1.030, "cement": 6.400}
    },
    "inspection chamber": {
        "short_desc": "I/C 90 x 45",
        "ratios": {"sand": 0.540, "rubble": 0.000, "brick": 527.00, "metal": 0.300, "cement": 3.530}
    },
    "m-10": {
        "short_desc": "P.C.C. 1:3:6",
        "ratios": {"sand": 0.470, "rubble": 0.000, "brick": 0.00, "metal": 0.940, "cement": 4.400}
    },
    "shahabad stone flooring": {
        "short_desc": "R/S Ladi",
        "ratios": {"sand": 0.022, "rubble": 0.000, "brick": 0.00, "metal": 0.000, "cement": 0.135}
    }
}