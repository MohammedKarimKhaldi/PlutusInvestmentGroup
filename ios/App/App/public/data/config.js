// Configuration file for dashboards (tracked in repo).
// You can edit this directly; the app will read it on page load.
// A legacy script exists at scripts/generate-config.js if you prefer to
// regenerate from an .env file, but it is not required for normal use.
window.DASHBOARD_CONFIG = {
    "dashboards": [
        {
            "id": "biolux",
            "name": "Biolux Analytics",
            "description": "Investor Network Live Dashboard",
            "excelUrl": "https://netorgft9359049-my.sharepoint.com/personal/mj_plutus-investment_com/_layouts/15/download.aspx?share=IQBgXPoLJxduQJ8V_2v17J27AZmmktQ4R_sp_0B4BqmxkMg",
            "sheets": {
                "funds": "Biolux Investors (Funds)",
                "familyOffices": "Biolux Investors (F.Os)",
                "figures": "figure"
            }
        },
        {
            "id": "wod",
            "name": "William Oak Diagnostics Analytics",
            "description": "Investor Network Live Dashboard",
            "excelUrl": "https://netorgft9359049-my.sharepoint.com/personal/es_plutus-investment_com/_layouts/15/download.aspx?share=IQB3aGDtg8uDS7_OITrGMMjAAdxoqzz-yY-LNaAer2dATBE",
            "sheets": {
                "funds": "WOD Investors (Funds)",
                "familyOffices": "WOD Investors (F.Os)",
                "figures": "figure"
            }
        },
        {
            "id": "IQ500",
            "name": "IQ500 Analytics",
            "description": "This is an example dashboard configuration.",
            "excelUrl": "https://netorgft9359049-my.sharepoint.com/personal/mj_plutus-investment_com/_layouts/15/download.aspx?share=IQB89WHDi04SQZTGfk5s0f5-AbnVhOGIQSHybJfNC5CDjtA",
            "sheets": {
                "funds": "IQ500 Investors (Funds)",
                "familyOffices": "IQ500 (F.Os)",
                "figures": "FiguresSheet"
            }
        },
        {
            "id": "adoram",
            "name": "Adoram Analytics",
            "description": "Investor Network Live Dashboard",
            "excelUrl": "https://netorgft9359049-my.sharepoint.com/personal/mj_plutus-investment_com/_layouts/15/download.aspx?share=IQDSdSA4sQQPQalStOoC8FQhAaFvnzhdPBjy6_oIJYyRPRs",
            "sheets": {
                "funds": "Adoram Investors (Funds)",
                "familyOffices": "Adoram Investors (F.Os)",
                "figures": "figure"
            }
        }
    ],
    "settings": {
        "defaultDashboard": "biolux",
        "allowLocalUpload": true,
        "title": "Investor Dashboard",
        "staffing": {
            "excelUrl": "https://netorgft9359049-my.sharepoint.com/:x:/r/personal/mj_plutus-investment_com/_layouts/15/download.aspx?share=IQCK2rboq1QyT6lCu-6uNxRrASgS1KRfEjISRy_YMAoiHIc",
            "sheetName": "Staffing"
        }
    }
};

window.DASHBOARD_PROXIES = [
    (url) => "https://api.codetabs.com/v1/proxy?quest=" + encodeURIComponent(url),
    (url) => "https://api.allorigins.win/get?url=" + encodeURIComponent(url)
];
