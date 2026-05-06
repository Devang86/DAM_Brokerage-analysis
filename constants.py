"""
Brokerage Data Analytics — Constants
KKC & Associates LLP
Branding, statutory rates, thresholds, audit procedures, column mappings.
"""

# ---------------------------------------------------------------------------
# Application metadata
# ---------------------------------------------------------------------------
APP_NAME = "Brokerage Data Analytics"
APP_VERSION = "2026.1"
APP_TAGLINE = "Stock Broker Audit Analytics"
APP_FULL_NAME = f"{APP_NAME} v{APP_VERSION}"

# ---------------------------------------------------------------------------
# KKC Brand
# ---------------------------------------------------------------------------
KKC_GREEN = "7CB542"
KKC_GREY = "808285"
WHITE = "FFFFFF"
LIGHT_GREEN = "E8F5E0"
LIGHT_GREEN_BG = "F0F7E8"
LIGHT_RED = "FDE8E8"
LIGHT_YELLOW = "FFF9E6"
LIGHT_BLUE = "F0F8FF"
LIGHT_GREY = "F2F2F2"
DARK_TEXT = "333333"
RED_ACCENT = "E74C3C"
AMBER_ACCENT = "F39C12"
BLUE_ACCENT = "3498DB"
FONT_NAME = "Source Sans Pro"

# ---------------------------------------------------------------------------
# Prescribed statutory rates (as of FY 2025-26)
# ---------------------------------------------------------------------------
STT_RATES = {
    "Cash Delivery Buy": 0.001,
    "Cash Delivery Sell": 0.001,
    "Cash Intraday Sell": 0.00025,
    "FO Futures Sell": 0.000125,
    "FO Options Sell": 0.000625,
}

SEBI_FEE_RATE = 10 / 10_000_000   # Rs 10 per crore = 0.000001

GST_RATE = 0.18   # 18% on brokerage

STAMP_DUTY_RATES = {
    "NSE_Cash_Buy": 0.00015,
    "BSE_Cash_Buy": 0.00015,
    "NSE_FO_Buy": 0.00003,
    "BSE_FO_Buy": 0.00003,
}

TURNOVER_CHARGE_RATES = {
    "NSE_Cash": 0.0000325,
    "BSE_Cash": 0.0000375,
    "NSE_FO_Futures": 0.0000190,
    "NSE_FO_Options": 0.0000500,
    "BSE_FO": 0.0000020,
}

# ---------------------------------------------------------------------------
# Default thresholds
# ---------------------------------------------------------------------------
DEFAULT_MATERIALITY = 500000
VARIANCE_FLAG_PCT = 20
CLIENT_CONCENTRATION_PCT = 5
ERROR_RATE_FLAG_PCT = 1
BROKERAGE_RATE_OUTLIER_STD = 2
ADJUSTMENT_FLAG_PCT = 5
MONTH_END_BUNCHING_PCT = 40
VOLUME_SPIKE_STD = 3
SEBI_MAX_BROKERAGE_PCT = 0.025   # 2.5% of turnover — SEBI prescribed maximum
CGST_RATE = 0.09                  # 9% of Net Brokerage (intra-state)
SGST_RATE = 0.09                  # 9% of Net Brokerage (intra-state)

# ---------------------------------------------------------------------------
# FY 2025-26 Exchange Holidays (NSE / BSE non-trading days)
# ---------------------------------------------------------------------------
FY2026_HOLIDAYS = [
    "2025-12-25",  # Christmas
    "2026-01-26",  # Republic Day
    "2026-03-03",  # Holi
    "2026-03-26",  # Shri Ram Navami
    "2026-03-31",  # Shri Mahavir Jayanti
    "2026-04-03",  # Good Friday
    "2026-04-14",  # Dr. Baba Saheb Ambedkar Jayanti
    "2026-05-01",  # Maharashtra Day
    "2026-05-28",  # Bakri Eid
    "2026-06-26",  # Moharram
    "2026-09-14",  # Ganesh Chaturthi
    "2026-10-02",  # Mahatma Gandhi Jayanti
    "2026-10-20",  # Dussehra
    "2026-11-10",  # Diwali — Balipratipada
    "2026-11-24",  # Prakash Gurpurb Sri Guru Nanak Dev
    "2026-12-25",  # Christmas
]

# ---------------------------------------------------------------------------
# Column name mappings (expected in trade-level RO sheets)
# ---------------------------------------------------------------------------
COL_CLIENT_CODE = "Client Code"
COL_CLIENT_NAME = "Client Name"
COL_SCRIP_CODE = "Scrip Code"
COL_SCRIP_NAME = "Scrip Name"
COL_EXCHANGE = "Exchange"
COL_TXN_DATE = "Txn Date"
COL_QTY = "Qty."
COL_BUY_SELL = "Buy / Sell"
COL_MARKET_VALUE = "Market Value"
COL_NET_VALUE = "Net Value"
COL_TOTAL_BROKERAGE = "Total Brokerage"
COL_STT = "STT"
COL_CGST = "CGST Charges"
COL_SGST = "SGST Charges"
COL_IGST = "IGST Charges"
COL_TURNOVER_TAX = "Turnover Tax"
COL_STAMP_CHARGES = "Stamp Charg."
COL_SEBI_FEES = "SEBI fees to be pass"
COL_NET_BROKERAGE = "Net Brokerage"
COL_TURNOVER = "Turnover"
COL_FAMILY_CODE = "Family Code"
COL_ISIN = "ISIN Code"
COL_CUSTODIAN = "Custodian"
COL_SETTLEMENT = "Settlement#"
COL_BOOK_TYPE = "BookType"
COL_INSTRUMENT = "Instrument"
COL_EXPIRY = "Expiry Date"
COL_STRIKE = "Strike Price"
COL_OPTION_TYPE = "OT"
COL_CONTRACT_VALUE = "Contract Value"
COL_SERVICE_TAX = "Service Tax"
COL_TOT_PASS = "TOT to be pass"
COL_BRANCH = "Branch Code"
COL_PRODUCT_DESC = "Product Description"

# Alternate column names (for flexible matching)
COL_ALIASES = {
    "Client Code": ["Client Code", "ClientCode", "client_code"],
    "Client Name": ["Client Name", "ClientName", "client_name"],
    "Txn Date": ["Txn Date", "Trade Date", "TxnDate", "Date"],
    "Total Brokerage": ["Total Brokerage", "Gross Brokerage", "Brokerage"],
    "Buy / Sell": ["Buy / Sell", "BuySell", "Buy/Sell", "Side", "P/S"],
    "Turnover": ["Turnover", "BE", "Market Value"],
    "STT": ["STT", "W"],
    "Stamp Charg.": ["Stamp Charg.", "Stamp Charges", "AH", "AW"],
    "SEBI fees to be pass": ["SEBI fees to be pass", "BH", "SEBI Fees"],
    "Product Description": ["Product Description", "ProductDesc", "Product Type", "Product"],
}

# ---------------------------------------------------------------------------
# Excel sheet configuration
# ---------------------------------------------------------------------------
EXCEL_SHEETS = [
    ("Cover", "Cover Page"),
    ("Executive Summary", "Key Metrics and Risk Flags"),
    ("Summary Dashboard", "Monthly Brokerage Trends"),
    ("Turnover Recon", "Turnover Reconciliation"),
    ("Client Concentration", "Client Concentration Analysis"),
    ("Scrip Analysis", "Scrip-wise Analysis"),
    ("Brokerage Rates", "Brokerage Rate Analytics"),
    ("Error Trades", "Error and Cancelled Trade Analysis"),
    ("GST Analytics", "GST Verification"),
    ("STT Verification", "STT Rate Verification"),
    ("SEBI Fees", "SEBI Fees Verification"),
    ("Stamp Duty", "Stamp Duty Analysis"),
    ("Adjustments", "Adjustment Analysis"),
    ("Buy Sell", "Buy and Sell Analysis"),
    ("Temporal", "Temporal Analysis"),
    ("Client Product", "Client × Product Analysis"),
    ("Txn Analysis", "Transaction-Level Analysis"),
    ("All Audit Procedures", "Master Audit Procedures"),
]

# ---------------------------------------------------------------------------
# Audit procedures (plain English for engagement team)
# Each section: list of (step_no, procedure, expected_evidence)
# ---------------------------------------------------------------------------
AUDIT_PROCEDURES = {
    "Summary Dashboard": [
        (1,
         "Reconcile gross brokerage to books: Add up the monthly gross brokerage from this report "
         "and agree it to the Profit & Loss account (Brokerage Income line). If there is a difference, "
         "trace it to the General Ledger and identify the missing entries.",
         "Signed P&L, General Ledger extract for Brokerage Income"),
        (2,
         "Compare to prior period: Obtain the prior year monthly brokerage data and compare month "
         "by month. Investigate any month where the difference is more than 20%. Ask management "
         "for the business reason (e.g., market rally, new clients, regulatory change).",
         "Prior year brokerage summary, management explanation letter"),
        (3,
         "Verify segment split: Check whether the NSE Cash, BSE Cash, and F&O split matches "
         "the segment-wise brokerage reports filed with the exchanges.",
         "Exchange-filed segment reports"),
        (4,
         "Assess net brokerage margin: Compute Net Brokerage / Gross Brokerage ratio. If it has "
         "changed significantly from the prior year, understand whether statutory cost rates changed "
         "or whether the broker is absorbing more costs.",
         "Prior year margin computation working"),
        (5,
         "Check MoM variance flags: For each month flagged with more than 20% variance, obtain "
         "management's explanation. Common reasons: market volatility, large client trades, "
         "regulatory changes. Document the explanation in working papers.",
         "Management representation for each flagged month"),
    ],
    "Turnover Recon": [
        (1,
         "Agree LD turnover to exchange statements: Obtain the exchange-issued turnover "
         "certificates (NSE and BSE) for each month. Agree the 'As per Exchange' column in "
         "this report to those certificates.",
         "Exchange-issued turnover certificates (NSE and BSE)"),
        (2,
         "Investigate all differences: For each month where the difference exceeds materiality, "
         "obtain a breakup of the difference from management. Common causes: interoperability "
         "trades, trade cancellations, timing differences.",
         "Management reconciliation schedule with breakup"),
        (3,
         "Understand interop adjustments: Under SEBI interoperability framework, a trade placed "
         "on BSE may be executed on NSE (or vice versa). This creates a reconciling difference. "
         "Verify the interop trades against the clearing corporation records.",
         "Clearing corporation interop trade report"),
        (4,
         "Document materiality assessment: Record the aggregate difference as a percentage of "
         "total turnover. If the unexplained difference exceeds performance materiality, "
         "consider it a potential misstatement.",
         "Materiality computation working paper"),
        (5,
         "Cross-check with SEBI filing: Verify that the turnover figures used for SEBI fee "
         "computation match the reconciled turnover.",
         "SEBI fee computation working, SEBI fee demand notices"),
    ],
    "Client Concentration": [
        (1,
         "Verify top client brokerage: For the top 10 clients by brokerage, select a sample of "
         "trades and verify the brokerage rate against the client agreement. Ensure no client "
         "is getting a rate below cost.",
         "Client agreements for top 10 clients, sample trade confirmations"),
        (2,
         "Check family code grouping: Verify that family codes are correctly assigned. Related "
         "clients (same family, same corporate group) should be grouped together. Incorrect "
         "grouping can hide concentration.",
         "Family code master list, client KYC records"),
        (3,
         "Assess dependency risk: If any single client contributes more than 5% of total "
         "brokerage, assess the business risk. Loss of that client would significantly impact "
         "the broker's revenue. Document this for the audit report.",
         "Revenue concentration analysis, going concern assessment"),
        (4,
         "Related party check: Cross-check the top 20 clients against the broker's related "
         "party list and promoter/director holdings. Trades by related parties must be "
         "disclosed separately under SEBI regulations.",
         "Related party register, promoter shareholding list"),
        (5,
         "HHI interpretation: If HHI exceeds 2500, the client base is highly concentrated. "
         "This means the broker's revenue depends on very few clients, which is a going "
         "concern risk indicator to flag in the audit report.",
         "HHI computation working, going concern checklist"),
        (6,
         "Rank top 20 clients by turnover and average brokerage rate: In addition to ranking "
         "by brokerage, identify the top 20 clients by turnover volume and by average brokerage "
         "rate charged. High-turnover clients with low rates may indicate preferential pricing; "
         "high-rate clients should be verified against agreements.",
         "Client ranking by turnover, client ranking by avg rate, rate agreement verification"),
        (7,
         "Month-on-month brokerage vs turnover analysis: Compare monthly brokerage and turnover "
         "trends side by side. If brokerage increases while turnover falls (or vice versa), "
         "investigate rate changes, client mix shifts, or booking errors.",
         "MoM brokerage vs turnover comparison, management explanation for divergences"),
    ],
    "Scrip Analysis": [
        (1,
         "Verify brokerage rates per scrip: For scrips with unusually high brokerage rates, "
         "check if a special arrangement exists (e.g., research-based premium brokerage). "
         "For unusually low rates, check if the rate covers the broker's costs.",
         "Scrip-wise rate analysis, special rate agreements"),
        (2,
         "Check for proprietary trading: Cross-check the top scrips with the broker's "
         "proprietary trading book. If the broker is also trading in the same scrips where "
         "it earns the most brokerage, there could be a conflict of interest.",
         "Proprietary trading book extract"),
        (3,
         "Verify scrip listing status: For scrips in the F&O segment, verify they are "
         "currently listed in the permitted derivatives list.",
         "Exchange derivative scrip list"),
        (4,
         "Cross-check with ISIN: Verify that the ISIN codes in the trade data match the "
         "scrip names. Mismatched ISINs could indicate data quality issues.",
         "ISIN master from depository"),
    ],
    "Brokerage Rates": [
        (1,
         "Verify against SEBI maximum limits: SEBI prescribes maximum brokerage rates. "
         "Verify no trade exceeds the maximum permitted rate. If it does, the excess "
         "brokerage must be refunded to the client.",
         "SEBI circular on maximum brokerage rates"),
        (2,
         "Check rate consistency with agreements: For the largest clients (top 10), "
         "obtain the brokerage rate agreement and verify that the actual rate charged "
         "matches the contractual rate.",
         "Client rate agreements, sample verification working"),
        (3,
         "Investigate rate outliers: For trades flagged as unusually high-rate, check if "
         "they are errors (e.g., data entry mistake) or genuine special-rate trades.",
         "Outlier investigation working, management explanation"),
        (4,
         "Assess rate trend: If average brokerage rates are declining month on month, "
         "understand whether this is due to competitive pressure, change in client mix, "
         "or shift towards low-margin F&O business.",
         "Monthly rate trend analysis, management discussion"),
        (5,
         "Segment-wise benchmark: Compare the broker's average rates by segment to "
         "industry benchmarks (available in SEBI annual reports).",
         "SEBI annual report, industry rate benchmarks"),
    ],
    "Error Trades": [
        (1,
         "Verify error trade reversal: For each cancelled or error trade, verify that a "
         "corresponding reversal entry exists in the books. The net impact of error plus "
         "reversal should be zero.",
         "Error trade listing, reversal journal entries"),
        (2,
         "Check authorization: Trade cancellation should be approved by a compliance "
         "officer or authorized person. Verify the approval trail for the largest error trades.",
         "Trade cancellation approval records, compliance officer sign-off"),
        (3,
         "Assess error rate: If the error rate exceeds 1% of total trades, it indicates "
         "a control weakness in the order management system. Discuss with management "
         "and assess the design effectiveness of controls over trade execution.",
         "Error rate computation, internal control questionnaire"),
        (4,
         "Investigate patterns: If the same client or same scrip repeatedly appears in "
         "error trades, investigate whether these are genuine errors or potential trade "
         "manipulation (e.g., wash trades, circular trading).",
         "Pattern analysis working, management explanation for repeated errors"),
        (5,
         "Quantify financial impact: Calculate the total brokerage, STT, and other "
         "charges on error trades. These should be fully reversed. Any residual amount "
         "is a misstatement.",
         "Financial impact computation working"),
    ],
    "GST Analytics": [
        (1,
         "Verify 18% rate application: GST on brokerage should be charged at 18% "
         "(9% CGST + 9% SGST for intra-state, or 18% IGST for inter-state). Check "
         "each flagged trade where the rate deviates.",
         "GST rate verification working for flagged trades"),
        (2,
         "Verify CGST/SGST vs IGST classification: IGST should be charged when the "
         "client is registered in a different state from the broker. Verify state "
         "classification using client PAN or registration details.",
         "Client state-wise classification list, PAN-state mapping"),
        (3,
         "Reconcile to GST returns: The total GST output in this analysis should "
         "reconcile to GSTR-1 (outward supplies). Obtain GSTR-1 summaries and compare.",
         "GSTR-1 monthly summaries, reconciliation working"),
        (4,
         "Check reverse charge applicability: Verify if any services received by the "
         "broker (e.g., from exchanges) attract reverse charge GST.",
         "Reverse charge liability computation, GSTR-3B returns"),
        (5,
         "Verify input credit: Cross-check GST paid on exchange transaction charges, "
         "SEBI fees, and other inputs against the broker's GSTR-3B return.",
         "GSTR-3B returns, ITC register"),
        (6,
         "CGST/SGST split verification: For all intra-state trades (where IGST is zero), "
         "verify that CGST is exactly 9% and SGST is exactly 9% of Net Brokerage. Flag "
         "any trade where either rate deviates beyond tolerance.",
         "CGST/SGST rate verification working for flagged trades"),
        (7,
         "IGST verification: For all inter-state trades (where CGST and SGST are zero), "
         "verify that IGST is exactly 18% of Net Brokerage. Flag deviations and verify "
         "the client's state classification is correct.",
         "IGST rate verification working, client state classification"),
    ],
    "STT Verification": [
        (1,
         "Verify STT rates by trade type: Cash delivery 0.1% on buy and sell. "
         "Cash intraday 0.025% on sell side only. F&O futures 0.0125% on sell side. "
         "F&O options 0.0625% on sell premium. Check the flagged trades.",
         "STT rate verification working for flagged trades"),
        (2,
         "Reconcile total STT to STT return: The total STT in this analysis should "
         "reconcile to the STT deposited (available in the broker's compliance records).",
         "STT deposit challans, compliance records"),
        (3,
         "Verify STT is charged to clients: STT is a pass-through. Ensure the full "
         "amount is recovered from clients and not absorbed by the broker.",
         "Client-wise STT recovery statement"),
        (4,
         "Check STT on error trades: For cancelled trades, verify that the STT "
         "charged has been reversed. Unreversed STT on cancelled trades is an "
         "overcharge to clients.",
         "Error trade STT reversal working"),
        (5,
         "Per-transaction STT rate flagging: For each cash delivery trade, verify that "
         "STT is levied at 0.1% of Market Value. Flag any trade where the effective STT "
         "rate deviates from the prescribed rate beyond tolerance.",
         "Per-transaction STT rate verification working"),
    ],
    "SEBI Fees": [
        (1,
         "Verify fee rate: SEBI charges turnover fees at prescribed rates "
         "(currently Rs 10 per crore). Compute the effective rate from this data "
         "and verify it matches the prescribed rate.",
         "SEBI fee rate circular, effective rate computation"),
        (2,
         "Reconcile to SEBI demand: Obtain the SEBI fee demand notices and compare "
         "with the computed fees. Any difference indicates either a computation "
         "error or a turnover reporting error.",
         "SEBI fee demand notices, reconciliation working"),
        (3,
         "Check exchange-wise computation: SEBI fees are collected by exchanges "
         "and remitted to SEBI. Verify that NSE and BSE are applying the same rate.",
         "Exchange-wise SEBI fee breakup"),
        (4,
         "Verify recovery from clients: Check whether SEBI fees are fully recovered "
         "from clients as part of the cost pass-through.",
         "Client-wise SEBI fee recovery statement"),
        (5,
         "Per-transaction SEBI fee verification: For each trade, verify that SEBI fees "
         "equal 0.0001% of Market Value (Rs 10 per crore). Flag any trade where the "
         "effective rate deviates from the prescribed rate beyond tolerance.",
         "Per-transaction SEBI fee rate verification working"),
    ],
    "Stamp Duty": [
        (1,
         "Verify stamp duty rates: Stamp duty on securities transactions is prescribed "
         "by the Indian Stamp Act (as amended by Finance Act 2019). Verify the rates "
         "applied match the prescribed rates for each instrument type.",
         "Indian Stamp Act schedule, rate verification working"),
        (2,
         "Reconcile to bank payments: Cross-check the monthly stamp duty computed "
         "from trade data against actual bank payments from the Stamp Duty Bank "
         "Payment file.",
         "Stamp duty bank payment statements, reconciliation working"),
        (3,
         "Check state-wise application: Stamp duty is now collected centrally and "
         "distributed to states. Verify the broker is applying the correct rate.",
         "Stamp duty rate schedule"),
        (4,
         "Recovery analysis: Stamp duty is a pass-through cost. Verify that the "
         "amount charged to clients matches the amount paid. Any shortfall is a "
         "cost to the broker; any excess is an overcharge to clients.",
         "Client-wise stamp duty recovery statement"),
    ],
    "Adjustments": [
        (1,
         "Understand each adjustment: For each monthly adjustment identified, obtain "
         "management's explanation. Common adjustments include: write-offs of "
         "uncollectable brokerage, promotional discounts, and error corrections.",
         "Adjustment register, management explanations"),
        (2,
         "Verify authorization: Adjustments to brokerage should be authorized by an "
         "appropriate person (e.g., compliance head, CFO). Check the approval trail.",
         "Adjustment approval records, authorization matrix"),
        (3,
         "Assess materiality: If total adjustments exceed 5% of gross brokerage, "
         "this is a material item. Consider whether it should be separately disclosed "
         "in the financial statements.",
         "Materiality computation, disclosure checklist"),
        (4,
         "Check for patterns: Repeated adjustments in the same direction (always "
         "reducing brokerage) may indicate systematic issues: either the initial "
         "brokerage computation is wrong, or adjustments are being used to manipulate "
         "reported revenue.",
         "Adjustment trend analysis, management discussion notes"),
        (5,
         "Year-end adjustments: Pay special attention to adjustments in March "
         "(year-end). Year-end adjustments may be used for revenue smoothing.",
         "March adjustment detail, cut-off testing working"),
    ],
    "Buy Sell": [
        (1,
         "Verify buy-sell balance: In an agency broking model, the broker should have "
         "roughly balanced buy and sell volumes. A large imbalance may indicate "
         "proprietary trading being mixed with client trades.",
         "Buy-sell analysis working, proprietary trade register"),
        (2,
         "Check large one-sided clients: Clients with heavily one-sided positions "
         "(all buys or all sells) could indicate: (a) genuine investment, "
         "(b) circular trading if paired with another client, or (c) money "
         "laundering indicators.",
         "Client position analysis, suspicious transaction report"),
        (3,
         "Delivery vs speculative classification: Where trades can be classified "
         "as delivery (held overnight) vs speculative (bought and sold same day), "
         "verify that the broker's internal classification matches.",
         "Trade classification report, settlement records"),
        (4,
         "Cross-check with depository: For net buy positions, verify that "
         "corresponding securities are credited to client demat accounts. "
         "For net sell positions, verify securities were available before sale.",
         "Depository participant statements, demat holding reports"),
    ],
    "Client Product": [
        (1,
         "Verify product-wise rate agreements: For top clients trading across multiple segments "
         "(NSE Cash, BSE Cash, F&O), obtain the rate agreement and verify that the brokerage "
         "rate charged per segment matches the contractual rate. Different segments typically "
         "have different agreed rates.",
         "Client rate agreements (segment-wise), sample verification working"),
        (2,
         "Assess cross-selling patterns: Identify clients trading in only one segment vs "
         "multiple segments. Understand whether the product mix is shifting towards lower-margin "
         "segments (e.g., F&O). A shift in mix can explain declining overall brokerage yields "
         "even without rate changes.",
         "Product mix analysis, prior year comparison, management discussion notes"),
        (3,
         "Check for segment rate arbitrage: Compare brokerage rates charged to the same client "
         "across segments. If a client receives significantly lower rates in one segment, "
         "investigate whether this is contractually justified or an error.",
         "Client-wise segment rate comparison, rate agreement verification"),
        (4,
         "Revenue mix risk assessment: If a large share of revenue comes from a single product "
         "segment, document this as a concentration risk. Regulatory changes affecting that "
         "segment (e.g., SEBI restrictions on F&O) could significantly impact revenue.",
         "Segment-wise revenue contribution, regulatory risk assessment"),
        (5,
         "Verify segment classification: Ensure trades are correctly classified into NSE Cash, "
         "BSE Cash, and F&O segments. Misclassification affects brokerage rate analysis, STT "
         "computation, and regulatory reporting.",
         "Segment-wise trade sample, exchange trade records"),
        (6,
         "Product type rate uniformity: For each client and product type (ALGO, ARBS, BKT, "
         "DEFAULT PRODUCT, ETF, OFS, BLOCK, BUYBACK), verify that the brokerage rate "
         "(Gross Brokerage / Market Value) is uniform across all trades. Non-uniform rates "
         "indicate either data errors or unapproved rate changes mid-period.",
         "Product type rate uniformity report, rate agreement verification"),
    ],
    "Transaction Analysis": [
        (1,
         "Verify high-value trades: For trades above materiality threshold, verify the trade "
         "details against contract notes issued to clients. Ensure trade price, quantity, "
         "brokerage rate, and statutory charges are correctly computed per SA 500.",
         "Contract notes for sampled trades, exchange trade confirmations"),
        (2,
         "Investigate zero brokerage trades: Trades with zero brokerage may indicate proprietary "
         "trades, promotional offers, or data errors. Verify each category and ensure proprietary "
         "trades are not mixed with client trades per SEBI regulations.",
         "Zero brokerage trade listing, management explanation, proprietary trade register"),
        (3,
         "Verify negative brokerage entries: Negative brokerage typically indicates trade "
         "reversals or credit notes. Each negative entry should have a corresponding original "
         "trade and proper authorization from compliance.",
         "Reversal authorization records, original trade matching working"),
        (4,
         "Audit sample testing: Select a sample of trades using an appropriate method (random, "
         "monetary unit, or stratified per SA 530). For each sampled trade, verify: "
         "(a) trade exists in exchange records, (b) brokerage rate per agreement, "
         "(c) statutory charges correctly computed, (d) contract note issued to client.",
         "Sampled trade verification working, exchange trade confirmations, contract notes"),
        (5,
         "Benford's Law analysis: Apply first-digit analysis to brokerage amounts. Significant "
         "deviation from expected Benford distribution may indicate data manipulation or "
         "systematic errors in brokerage computation. Use MAD (Mean Absolute Deviation) test.",
         "Benford's Law test results, investigation working for deviations"),
        (6,
         "SEBI maximum brokerage check: Verify no trade exceeds the SEBI prescribed maximum "
         "brokerage rate of 2.5% of transaction value. Any excess must be refunded to clients "
         "and reported to compliance.",
         "SEBI max brokerage verification working, refund computation"),
        (7,
         "High turnover with low brokerage: Identify trades in the top 25% by turnover "
         "but bottom 10% by brokerage rate. These may indicate preferential pricing, "
         "revenue leakage, or promotional arrangements requiring management explanation.",
         "High turnover / low brokerage trade listing, management explanation"),
    ],
    "Temporal": [
        (1,
         "Investigate volume spikes: For days flagged with abnormally high trading "
         "volume, check whether they coincide with market events (budget, RBI policy, "
         "corporate results). Spikes without a market event may indicate manipulative "
         "trading.",
         "Volume spike dates vs market event calendar"),
        (2,
         "Check month-end bunching: If trading activity is concentrated in the last "
         "few days of the month, investigate whether clients are being encouraged to "
         "trade to meet the broker's revenue targets.",
         "Month-end trade distribution analysis"),
        (3,
         "Day-of-week patterns: Stock markets are closed on weekends. If any trades "
         "show Saturday or Sunday dates, these are data errors that need correction.",
         "Weekend trade listing (should be nil)"),
        (4,
         "Settlement-wise analysis: Verify that settlement obligations (pay-in of "
         "securities and funds) are being met on T+1 basis as per current SEBI norms.",
         "Settlement obligation reports from clearing corporation"),
        (5,
         "Assess seasonal pattern: Understand whether the broker's business has "
         "seasonal patterns (e.g., higher volumes during derivative expiry weeks, "
         "budget month, IPO seasons). This helps form audit expectations per SA 520.",
         "Monthly volume analysis, SA 520 analytical procedures working"),
        (6,
         "Exchange holiday detection: Verify that no trades have been executed on "
         "exchange holidays (as per the NSE/BSE holiday calendar for FY 2025-26). "
         "Trades on holidays indicate data errors requiring investigation.",
         "Exchange holiday calendar, holiday trade listing (should be nil)"),
        (7,
         "Daily turnover ranking: Rank all trading days by total turnover from highest "
         "to lowest. Select the top trading days for obtaining and verifying the OTR "
         "(Order Trade Register) files on a sample basis per SA 500/530.",
         "Daily turnover ranking, OTR files for selected high-value days"),
    ],
}
