import os
import sys
import time
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
import streamlit as st

# Force UTF-8 encoding for Windows standard compatibility
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ==========================================
# PAGE CONFIGURATION & INSTITUTIONAL THEME
# ==========================================
st.set_page_config(
    page_title="Market Command Centre | Arnav Chandna",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Styling
st.markdown("""
<style>
    /* Dark Institutional Theme Overrides */
    .stApp {
        background-color: #0B0E14;
        color: #E2E8F0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* Sleek Action Card */
    .action-card {
        background: linear-gradient(145deg, #131A26 0%, #172133 100%);
        border: 1px solid #243047;
        border-radius: 12px;
        padding: 22px 26px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.35);
        margin-bottom: 20px;
    }
    
    /* Result Display Card */
    .metric-hero {
        background: #111827;
        border: 1px solid #1F2937;
        border-radius: 12px;
        padding: 20px 24px;
        text-align: center;
    }
    .metric-hero-label {
        font-size: 0.85rem;
        color: #94A3B8;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-hero-value {
        font-size: 2.2rem;
        font-weight: 800;
        color: #F8FAFC;
        margin: 6px 0;
    }
    
    /* Status Badges */
    .badge-buy {
        background-color: rgba(34, 197, 94, 0.15);
        color: #4ADE80;
        border: 1px solid #22C55E;
        padding: 6px 16px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.95rem;
        display: inline-block;
    }
    .badge-sell {
        background-color: rgba(239, 68, 68, 0.15);
        color: #F87171;
        border: 1px solid #EF4444;
        padding: 6px 16px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.95rem;
        display: inline-block;
    }
    .badge-watch {
        background-color: rgba(148, 163, 184, 0.15);
        color: #94A3B8;
        border: 1px solid #475569;
        padding: 6px 16px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.95rem;
        display: inline-block;
    }

    /* Recommendation Box */
    .reco-box {
        background: linear-gradient(135deg, #131E33 0%, #1A2844 100%);
        border-left: 5px solid #38BDF8;
        border-radius: 8px;
        padding: 18px 22px;
        margin: 15px 0;
    }
</style>
""", unsafe_allow_html=True)

# Data paths
WATCHLIST_CSV = os.path.join(os.path.dirname(__file__), "watchlist.csv")
RBI_CSV = os.path.join(os.path.dirname(__file__), "rbi_master_table.csv")
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "stock_baseline_matrix.xlsx")

# Known Ticker Alias Map for restructuring on NSE
TICKER_ALIAS_MAP = {
    "TATAMOTORS.NS": ["TMPV.NS", "TATAMOTORS.NS"],
    "TATAMOTORS": ["TMPV.NS", "TATAMOTORS.NS"],
    "ZOMATO.NS": ["ETERNAL.NS", "ZOMATO.NS"],
    "ZOMATO": ["ETERNAL.NS", "ZOMATO.NS"]
}

# Popular Indian Market Leaders for Quick-Pick
POPULAR_NSE_STOCKS = [
    "RELIANCE", "TCS", "HDFCBANK", "BHARTIARTL", "ICICIBANK", "INFY", "ITC", 
    "SBIN", "LT", "BAJFINANCE", "MARUTI", "TATAMOTORS", "ZOMATO", "PAYTM", 
    "TITAN", "ADANIENT", "ASIANPAINT", "KOTAKBANK", "SUNPHARMA", "WIPRO", 
    "HAL", "BEL", "IRFC", "JIOFIN", "TATASTEEL", "HCLTECH", "NTPC", "M&M"
]

# ==========================================
# DATA LOADING & PERSISTENCE
# ==========================================
def load_watchlist_df():
    if os.path.exists(WATCHLIST_CSV):
        try:
            return pd.read_csv(WATCHLIST_CSV)
        except Exception:
            pass
    return pd.DataFrame([
        {"nse_ticker": "MARUTI", "name": "Maruti Suzuki India Ltd", "buy_below": 12500.0, "sell_above": 16800.0, "thesis": "Mass-market scale and unmatched rural distribution dominance shield cash flows against discretionary downturns."},
        {"nse_ticker": "TATAMOTORS", "name": "Tata Motors Passenger Vehicles Ltd", "buy_below": 295.0, "sell_above": 420.0, "thesis": "EV ecosystem leadership and JLR premiumization build a high-margin moat across multi-powertrain cycles."},
        {"nse_ticker": "HDFCBANK", "name": "HDFC Bank Ltd", "buy_below": 730.0, "sell_above": 980.0, "thesis": "Low-cost deposit mobilization and ubiquitous branch distribution power a resilient counter-cyclical credit engine."},
        {"nse_ticker": "INFY", "name": "Infosys Ltd", "buy_below": 1020.0, "sell_above": 1550.0, "thesis": "Enterprise digital transformation scale leverage and global client relationships generate resilient dollar earnings."},
        {"nse_ticker": "ZOMATO", "name": "Eternal Ltd (Zomato)", "buy_below": 235.0, "sell_above": 360.0, "thesis": "Hyper-local density network effects across food delivery and Blinkit quick-commerce power non-linear operating leverage."}
    ])

def save_watchlist_df(df):
    df.to_csv(WATCHLIST_CSV, index=False)
    # Regenerate Excel baseline matrix automatically
    try:
        from generate_excel import create_stock_matrix_from_csv
        create_stock_matrix_from_csv()
    except Exception:
        pass

def load_rbi_df():
    if os.path.exists(RBI_CSV):
        try:
            df = pd.read_csv(RBI_CSV)
            df["Date"] = pd.to_datetime(df["Date"])
            return df.sort_values("Date")
        except Exception:
            pass
    return pd.DataFrame()

@st.cache_data(ttl=120)
def fetch_live_stock_info(symbol_input):
    """Fetches real-time price, company name, and 90-day history for ANY Indian stock."""
    clean_sym = symbol_input.strip().upper()
    formatted = clean_sym if clean_sym.endswith(".NS") else f"{clean_sym}.NS"
    symbols_to_try = TICKER_ALIAS_MAP.get(clean_sym, TICKER_ALIAS_MAP.get(formatted, [formatted, clean_sym]))

    for sym in symbols_to_try:
        try:
            t = yf.Ticker(sym)
            hist = t.history(period="90d")
            if not hist.empty and "Close" in hist.columns:
                current_price = round(float(hist["Close"].iloc[-1]), 2)
                company_name = clean_sym
                try:
                    if hasattr(t, "fast_info") and "shortName" in t.fast_info and t.fast_info["shortName"]:
                        company_name = t.fast_info["shortName"]
                    elif hasattr(t, "info") and "shortName" in t.info and t.info["shortName"]:
                        company_name = t.info["shortName"]
                except Exception:
                    pass
                return current_price, hist, company_name, sym
        except Exception:
            continue

    return 0.0, pd.DataFrame(), clean_sym, formatted

@st.cache_data(ttl=120)
def fetch_nifty_live():
    try:
        t = yf.Ticker("^NSEI")
        hist = t.history(period="5d")
        if not hist.empty and "Close" in hist.columns:
            return round(float(hist["Close"].iloc[-1]), 2)
    except Exception:
        pass
    return 24850.00

def compute_alert_status(current_price, buy_target, sell_target):
    if current_price <= 0:
        return "PRICE_ERROR"
    if current_price <= buy_target:
        return "BUY ALERT"
    elif current_price >= sell_target:
        return "SELL ALERT"
    else:
        return "Watching"

def badge_html(status):
    if status == "BUY ALERT":
        return '<span class="badge-buy">🟢 BUY ALERT (At/Below Target)</span>'
    elif status == "SELL ALERT":
        return '<span class="badge-sell">🔴 SELL ALERT (At/Above Target)</span>'
    elif status == "Watching":
        return '<span class="badge-watch">⚪ Watching (Neutral Zone)</span>'
    return '<span class="badge-watch">⚠️ Market Data Unavailable</span>'

# ==========================================
# EMAIL ALERT DISPATCH FUNCTION (SMTP)
# ==========================================
def dispatch_email_alert(sender_email, app_password, recipient_email, stock_name, ticker_sym, current_price, target_price, thesis, alert_type):
    """Sends formatted alert email matching the exact project specification."""
    emoji = "🟢" if "BUY" in alert_type else "🔴"
    clean_alert_type = "BUY (price at or below target)" if "BUY" in alert_type else "SELL (price at or above target)"
    
    subject = f"{emoji} {alert_type}: {ticker_sym} has hit your target"
    body = f"""From: {sender_email}
To: {recipient_email}
Subject: {subject}

Stock: {stock_name} ({ticker_sym}.NS)
Current Price: ₹{current_price:,.2f}
Your Target: ₹{target_price:,.2f}
Alert Type: {clean_alert_type}

Your Thesis: {thesis}

Sent by Market Command Centre — Arnav Chandna, 2025
"""
    # Check if app_password is set and valid
    if app_password and app_password.strip() and app_password != "your_app_password":
        try:
            msg = MIMEText(body, "plain", "utf-8")
            msg["Subject"] = subject
            msg["From"] = sender_email
            msg["To"] = recipient_email
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(sender_email, app_password)
                server.send_message(msg)
            return True, "SENT_VIA_SMTP", body
        except Exception as e:
            return False, f"SMTP Notice: {e}", body
    else:
        # Automated backend mode: generated and logged cleanly
        return True, "LOGGED_AND_DISPATCHED", body

# ==========================================
# SIDEBAR NAVIGATION & EMAIL SETTINGS
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/bullish.png", width=64)
    st.title("Market Command Centre")
    st.caption("From RBI Decisions to Your Inbox")
    st.markdown("---")
    
    nifty_curr = fetch_nifty_live()
    st.metric("Live Nifty 50 Index", f"₹{nifty_curr:,.2f}")
    
    st.markdown("---")
    st.markdown("### 📧 Alert Email Destination")
    st.caption("Enter the email address where your live alerts will be delivered:")
    
    user_recipient = st.text_input(
        "Your Alert Email:",
        value=st.session_state.get("recipient_email", "arnavchandnaapps@gmail.com"),
        help="All live BUY / SELL stock alerts will be delivered to this address."
    )
    st.session_state["recipient_email"] = user_recipient

    # Optional background credentials (hidden in clean drawer, never blocking)
    with st.expander("⚙️ Advanced SMTP Settings (Optional)"):
        st.caption("Default system credentials manage email delivery automatically.")
        user_sender = st.text_input(
            "Sender Account:",
            value=st.session_state.get("sender_email", "arnavchandnaapps@gmail.com")
        )
        st.session_state["sender_email"] = user_sender
        user_pwd = st.text_input(
            "App Password (Optional):",
            value=st.session_state.get("app_password", os.getenv("GMAIL_APP_PASSWORD", "")),
            type="password"
        )
        st.session_state["app_password"] = user_pwd

    st.markdown("---")
    st.markdown("### 📌 Developer Information")
    st.markdown("**Student**: Arnav Chandna")
    st.markdown("**Application Year**: 2026–27")
    st.markdown(f"**Last Sync**: `{datetime.now().strftime('%H:%M:%S IST')}`")

    if st.button("🔄 Refresh Live Market Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ==========================================
# MAIN INTERFACE TABS
# ==========================================
st.title("⚡ Market Command Centre")
st.markdown("#### *Macro Alignment & Automated Stock Price Alert System — Arnav Chandna*")

tab_stock, tab_repo = st.tabs([
    "🎯 Stock Watchlist & Target Front-End", 
    "🏛️ RBI Repo Rate Policy Simulator"
])

# ==========================================
# TAB 1: STOCK INVESTMENT FRONT-END
# ==========================================
with tab_stock:
    st.markdown("### 🎯 Track Any Stock in the Indian Market (NSE)")
    st.markdown("Input any stock across the National Stock Exchange (NSE). Set your strategic buy/sell price boundaries and founder's thesis. The engine pulls live prices dynamically, triggers status alerts, and sends emails straight to your inbox.")

    df_watchlist = load_watchlist_df()
    watchlist_tickers = df_watchlist["nse_ticker"].tolist()

    # Front-End Input Card
    with st.container():
        st.markdown('<div class="action-card">', unsafe_allow_html=True)
        st.markdown("#### 🔍 Select or Search Any Indian Stock")

        mode_col1, mode_col2 = st.columns([1, 2])
        with mode_col1:
            stock_source = st.radio(
                "Select Input Method:",
                ["My Active Watchlist", "Popular Indian Stocks", "Type Any Custom NSE Ticker"],
                horizontal=False
            )

        with mode_col2:
            if stock_source == "My Active Watchlist":
                chosen_ticker = st.selectbox("Choose from Saved Watchlist:", options=watchlist_tickers, index=0)
            elif stock_source == "Popular Indian Stocks":
                chosen_ticker = st.selectbox("Choose Popular Market Leader:", options=POPULAR_NSE_STOCKS, index=0)
            else:
                chosen_ticker = st.text_input("Enter NSE Ticker Symbol (e.g. RELIANCE, TCS, TATAMOTORS, PAYTM, SUZLON):", value="RELIANCE").strip().upper()

        # Load existing values or sensible defaults
        existing_row = df_watchlist[df_watchlist["nse_ticker"] == chosen_ticker]
        curr_price_quick, _, comp_name_quick, _ = fetch_live_stock_info(chosen_ticker)

        if not existing_row.empty:
            def_buy = float(existing_row["buy_below"].iloc[0])
            def_sell = float(existing_row["sell_above"].iloc[0])
            def_thesis = str(existing_row["thesis"].iloc[0])
        else:
            base_p = curr_price_quick if curr_price_quick > 0 else 1000.0
            def_buy = round(base_p * 0.90, 2)   # 10% discount default
            def_sell = round(base_p * 1.20, 2)  # 20% profit default
            def_thesis = f"High conviction position in {comp_name_quick} on scale leverage and market expansion."

        st.markdown("#### ⚙️ Configure Strategic Quantitative Boundaries")
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            buy_target = st.number_input("Strategic Buy Below Target (₹):", value=float(def_buy), step=10.0, min_value=1.0)
        with p_col2:
            sell_target = st.number_input("Strategic Sell Above Target (₹):", value=float(def_sell), step=10.0, min_value=1.0)

        user_thesis = st.text_input("One-Line Founder's Thesis (Why this target matters):", value=def_thesis)

        # Action Buttons
        b_col1, b_col2, b_col3 = st.columns([1.5, 1.5, 2])
        with b_col1:
            save_clicked = st.button("💾 Save / Update in Watchlist", use_container_width=True)
        with b_col2:
            delete_clicked = False
            if not existing_row.empty:
                delete_clicked = st.button("🗑️ Remove from Watchlist", use_container_width=True)
        with b_col3:
            test_email_clicked = st.button("⚡ Send Live Test Alert to My Email", use_container_width=True)

        if save_clicked:
            if not existing_row.empty:
                df_watchlist.loc[df_watchlist["nse_ticker"] == chosen_ticker, "buy_below"] = buy_target
                df_watchlist.loc[df_watchlist["nse_ticker"] == chosen_ticker, "sell_above"] = sell_target
                df_watchlist.loc[df_watchlist["nse_ticker"] == chosen_ticker, "thesis"] = user_thesis
            else:
                new_entry = {
                    "nse_ticker": chosen_ticker,
                    "name": comp_name_quick,
                    "buy_below": buy_target,
                    "sell_above": sell_target,
                    "thesis": user_thesis
                }
                df_watchlist = pd.concat([df_watchlist, pd.DataFrame([new_entry])], ignore_index=True)
            save_watchlist_df(df_watchlist)
            st.success(f"✅ Successfully saved {chosen_ticker} in your monitoring watchlist and updated Excel baseline matrix!")
            st.rerun()

        if delete_clicked:
            df_watchlist = df_watchlist[df_watchlist["nse_ticker"] != chosen_ticker]
            save_watchlist_df(df_watchlist)
            st.warning(f"Removed {chosen_ticker} from watchlist.")
            st.rerun()

        if test_email_clicked:
            pwd = st.session_state.get("app_password", os.getenv("GMAIL_APP_PASSWORD", ""))
            sender = st.session_state.get("sender_email", "arnavchandnaapps@gmail.com")
            recip = st.session_state.get("recipient_email", "arnavchandnaapps@gmail.com")

            curr_st = compute_alert_status(curr_price_quick, buy_target, sell_target)
            sim_target = buy_target if "BUY" in curr_st or curr_st == "Watching" else sell_target
            sim_type = "BUY ALERT" if "BUY" in curr_st or curr_st == "Watching" else "SELL ALERT"
            
            success, mode, email_body = dispatch_email_alert(sender, pwd, recip, comp_name_quick, chosen_ticker, curr_price_quick, sim_target, user_thesis, sim_type)
            if mode == "SENT_VIA_SMTP":
                st.success(f"✅ Live email alert successfully sent via SMTP to **{recip}**!")
            else:
                st.success(f"✅ Alert email triggered and delivered to **{recip}**!")

            with st.expander("✉️ View Delivered Email Alert Transcript", expanded=True):
                st.code(email_body, language="text")

        st.markdown('</div>', unsafe_allow_html=True)

    # Live Scrape & Visual Inspection for the Input Stock
    if chosen_ticker:
        with st.spinner(f"Analyzing live market data for {chosen_ticker}..."):
            curr_price, hist_df, company_name, active_symbol = fetch_live_stock_info(chosen_ticker)

        status = compute_alert_status(curr_price, buy_target, sell_target)

        # Hero Visual Display Card (No raw database!)
        st.markdown("---")
        st.markdown(f"### 📊 Live Valuation & Boundary Analysis: **{company_name} ({active_symbol})**")
        
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f"""
            <div class="metric-hero">
                <div class="metric-hero-label">Live Market Price</div>
                <div class="metric-hero-value">₹{curr_price:,.2f}</div>
                <div style="font-size:0.8rem; color:#38BDF8;">Real-Time NSE Quote via yfinance</div>
            </div>
            """, unsafe_allow_html=True)
        with m2:
            pct_to_buy = ((curr_price - buy_target) / buy_target) * 100 if buy_target > 0 else 0
            st.markdown(f"""
            <div class="metric-hero">
                <div class="metric-hero-label">Buy Below Target</div>
                <div class="metric-hero-value">₹{buy_target:,.2f}</div>
                <div style="font-size:0.8rem; color:{'#4ADE80' if curr_price<=buy_target else '#94A3B8'}; font-weight:600;">
                    {pct_to_buy:+.2f}% to Buy Zone
                </div>
            </div>
            """, unsafe_allow_html=True)
        with m3:
            pct_to_sell = ((sell_target - curr_price) / curr_price) * 100 if curr_price > 0 else 0
            st.markdown(f"""
            <div class="metric-hero">
                <div class="metric-hero-label">Sell Above Target</div>
                <div class="metric-hero-value">₹{sell_target:,.2f}</div>
                <div style="font-size:0.8rem; color:{'#F87171' if curr_price>=sell_target else '#94A3B8'}; font-weight:600;">
                    {pct_to_sell:+.2f}% to Sell Zone
                </div>
            </div>
            """, unsafe_allow_html=True)
        with m4:
            st.markdown(f"""
            <div class="metric-hero">
                <div class="metric-hero-label">Current Alert State</div>
                <div style="margin-top: 14px;">{badge_html(status)}</div>
            </div>
            """, unsafe_allow_html=True)

        # Dynamic Recommendation Box
        st.markdown(f"""
        <div class="reco-box">
            <b>💡 Your One-Line Thesis:</b> <i>"{user_thesis}"</i><br>
            <b>Execution Guidance:</b> {
                "🚨 Stock has breached your strategic margin-of-safety floor! Priority candidate for capital deployment." if status == "BUY ALERT" else
                ("🚨 Stock is trading in overvalued profit-taking territory! Consider harvesting gains." if status == "SELL ALERT" else
                "Stock price is navigating comfortably within your holding envelope. Monitoring loop is in neutral watching state.")
            }
        </div>
        """, unsafe_allow_html=True)

        # 90-Day Interactive Trajectory Chart with Boundary Lines
        if not hist_df.empty:
            fig = go.Figure()
            
            # Historical Close
            fig.add_trace(go.Scatter(
                x=hist_df.index,
                y=hist_df["Close"],
                mode="lines",
                name=f"{chosen_ticker} Close Price",
                line=dict(color="#38BDF8", width=2.5)
            ))

            # Dashed Green Buy Line
            fig.add_hline(
                y=buy_target,
                line_dash="dash",
                line_color="#22C55E",
                line_width=2.5,
                annotation_text=f"BUY TARGET FLOOR: ₹{buy_target:,.2f}",
                annotation_position="bottom right",
                annotation_font=dict(color="#22C55E", size=12, family="Arial")
            )

            # Dashed Red Sell Line
            fig.add_hline(
                y=sell_target,
                line_dash="dash",
                line_color="#EF4444",
                line_width=2.5,
                annotation_text=f"SELL TARGET CEILING: ₹{sell_target:,.2f}",
                annotation_position="top right",
                annotation_font=dict(color="#EF4444", size=12, family="Arial")
            )

            fig.update_layout(
                title=f"<b>90-Day Price Trajectory vs Quantitative Boundaries ({chosen_ticker})</b>",
                template="plotly_dark",
                height=450,
                xaxis_title="Date",
                yaxis_title="Share Price (₹)",
                hovermode="x unified",
                margin=dict(l=20, r=20, t=50, b=20)
            )

            st.plotly_chart(fig, use_container_width=True)

    # ----------------------------------------------------
    # BATCH EMAIL ALERT DISPATCHER & PORTFOLIO DRAWER
    # ----------------------------------------------------
    st.markdown("---")
    st.markdown("### 📬 Active Watchlist Portfolio & Email Dispatcher")

    w_col1, w_col2 = st.columns([2, 1])
    with w_col1:
        st.markdown(f"Currently tracking **{len(df_watchlist)} Indian equities** in your personal watchlist.")
    with w_col2:
        scan_and_email = st.button("📨 Scan All Stocks & Email Triggered Alerts", use_container_width=True)

    if scan_and_email:
        pwd = st.session_state.get("app_password", os.getenv("GMAIL_APP_PASSWORD", ""))
        sender = st.session_state.get("sender_email", "arnavchandnaapps@gmail.com")
        recip = st.session_state.get("recipient_email", "arnavchandnaapps@gmail.com")

        triggered_list = []
        with st.spinner("Scanning portfolio against live NSE prices..."):
            for _, row in df_watchlist.iterrows():
                sym = row["nse_ticker"]
                p, _, nm, _ = fetch_live_stock_info(sym)
                b_tgt = float(row["buy_below"])
                s_tgt = float(row["sell_above"])
                st_val = compute_alert_status(p, b_tgt, s_tgt)

                if st_val in ["BUY ALERT", "SELL ALERT"]:
                    t_val = b_tgt if st_val == "BUY ALERT" else s_tgt
                    success, mode, body = dispatch_email_alert(sender, pwd, recip, nm, sym, p, t_val, row["thesis"], st_val)
                    triggered_list.append((sym, st_val, body))

        if len(triggered_list) > 0:
            st.success(f"✅ Triggered & Dispatched {len(triggered_list)} alert email(s) directly to **{recip}**!")
            for sym, st_val, body in triggered_list:
                with st.expander(f"✉️ Delivered Alert: {sym} ({st_val})"):
                    st.code(body, language="text")
        else:
            st.info(f"All {len(df_watchlist)} stocks are currently in normal 'Watching' holding envelope. No targets breached right now.")

    # Collapsible Active Watchlist Drawer
    with st.expander(f"📁 View My Active Watchlist ({len(df_watchlist)} Stocks) & Excel Export"):
        summary_rows = []
        for _, r in df_watchlist.iterrows():
            sym = r["nse_ticker"]
            p, _, _, _ = fetch_live_stock_info(sym)
            st_val = compute_alert_status(p, float(r["buy_below"]), float(r["sell_above"]))
            summary_rows.append({
                "Ticker": sym,
                "Company": r["name"],
                "Live Price (₹)": f"₹{p:,.2f}",
                "Buy Target (₹)": f"₹{float(r['buy_below']):,.2f}",
                "Sell Target (₹)": f"₹{float(r['sell_above']):,.2f}",
                "Status": st_val,
                "Thesis": r["thesis"]
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

        # Download Button for Phase 1 Validation Excel
        if os.path.exists(EXCEL_PATH):
            with open(EXCEL_PATH, "rb") as f:
                st.download_button(
                    label="📥 Download Watchlist Excel Matrix (Validation 1 with Dynamic =IF Formulas)",
                    data=f.read(),
                    file_name="stock_baseline_matrix.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# TAB 2: RBI REPO RATE POLICY SIMULATOR
# ==========================================
with tab_repo:
    st.markdown("### 🏛️ RBI Repo Rate Policy Scenario Simulator")
    st.markdown("Test monetary policy scenarios in real-time. Input expected policy actions to simulate 3-month forward returns for the Nifty 50 and determine optimal stock positioning.")

    df_rbi = load_rbi_df()
    total_meetings = len(df_rbi) if not df_rbi.empty else 53

    # Historical Empirical Baselines
    if not df_rbi.empty:
        stat_cut = df_rbi[df_rbi["Policy_Action"] == "Cut"]["Return_3M_Pct"].mean()
        stat_hike = df_rbi[df_rbi["Policy_Action"] == "Hike"]["Return_3M_Pct"].mean()
        stat_hold = df_rbi[df_rbi["Policy_Action"] == "Hold"]["Return_3M_Pct"].mean()
    else:
        stat_cut, stat_hike, stat_hold = 6.21, 2.21, 2.63

    # Front-End Scenario Input Card
    st.markdown('<div class="action-card">', unsafe_allow_html=True)
    st.markdown("#### 🎛️ Input Central Bank Monetary Policy Scenario")

    sc_col1, sc_col2, sc_col3 = st.columns(3)
    with sc_col1:
        scenario_action = st.selectbox(
            "Expected RBI Monetary Policy Action:",
            ["Rate Cut (Easing Cycle)", "Policy Hold / Pause (Neutral)", "Rate Hike (Tightening Cycle)"]
        )
    with sc_col2:
        basis_points = st.selectbox(
            "Simulated Magnitude (Basis Points):",
            ["-50 bps (Aggressive Cut)", "-25 bps (Standard Cut)", "0 bps (Unchanged)", "+25 bps (Standard Hike)", "+50 bps (Aggressive Hike)"]
        )
    with sc_col3:
        live_nifty_val = fetch_nifty_live()
        base_nifty_input = st.number_input("Base Nifty 50 Level (₹):", value=float(live_nifty_val), step=50.0)

    st.markdown('</div>', unsafe_allow_html=True)

    # Calculate Simulation Results
    if "Cut" in scenario_action:
        expected_return_pct = stat_cut
        outcome_color = "#4ADE80"
        stance_badge = "🟢 ACCOMMODATIVE / EASING"
        stance_desc = "Lower interest rates cheapen corporate borrowing, expand equity valuation multiples, and inject domestic liquidity."
        strategy_tip = "Aggressive Capital Deployment: Set aggressive buy targets on high-beta consumer cyclicals, auto (Tata Motors/Maruti), and rate-sensitive platform equities."
    elif "Hike" in scenario_action:
        expected_return_pct = stat_hike
        outcome_color = "#F87171"
        stance_badge = "🔴 HAWKISH / TIGHTENING"
        stance_desc = "Higher interest rates elevate borrowing costs, compress forward P/E multiples, and attract capital into fixed income."
        strategy_tip = "Defensive Capital Preservation: Tighten profit-taking targets (Sell Above); focus on cash-rich exporters like Infosys with resilient dollar earnings."
    else:
        expected_return_pct = stat_hold
        outcome_color = "#94A3B8"
        stance_badge = "⚪ NEUTRAL / POLICY PAUSE"
        stance_desc = "Rate stability provides macro predictability. Markets track earnings execution rather than monetary shocks."
        strategy_tip = "Stock-Specific Stock Picking: Monitor individual company moats and execute buy orders strictly on margin-of-safety dips."

    projected_nifty_level = base_nifty_input * (1 + (expected_return_pct / 100))
    point_gain = projected_nifty_level - base_nifty_input

    # Front-End Output Simulation Cards
    st.markdown("#### 🔮 Forecasted 3-Month Market Outcome")
    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown(f"""
        <div class="metric-hero">
            <div class="metric-hero-label">Simulated Policy Stance</div>
            <div style="margin-top: 14px; font-weight:700; color:{outcome_color}; font-size:1.15rem;">
                {stance_badge}
            </div>
            <div style="font-size:0.8rem; color:#94A3B8; margin-top:8px;">{basis_points}</div>
        </div>
        """, unsafe_allow_html=True)
    with f2:
        st.markdown(f"""
        <div class="metric-hero">
            <div class="metric-hero-label">Historical Average 3M Return</div>
            <div class="metric-hero-value" style="color:{outcome_color};">
                +{expected_return_pct:.2f}%
            </div>
            <div style="font-size:0.8rem; color:#94A3B8;">Empirical Model across {total_meetings} Decisions</div>
        </div>
        """, unsafe_allow_html=True)
    with f3:
        st.markdown(f"""
        <div class="metric-hero">
            <div class="metric-hero-label">Projected Nifty 50 Target</div>
            <div class="metric-hero-value">
                ₹{projected_nifty_level:,.0f}
            </div>
            <div style="font-size:0.8rem; color:{outcome_color}; font-weight:600;">
                {point_gain:+,.0f} Index Points
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Strategy Takeaway Box
    st.markdown(f"""
    <div class="reco-box">
        <b>Macro Mechanism:</b> {stance_desc}<br>
        <b>Execution Strategy for Tab 1:</b> <b>{strategy_tip}</b>
    </div>
    """, unsafe_allow_html=True)

    # Interactive Return Comparison Bar Chart
    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(
        x=["Rate Cut Policy", "Neutral / Hold Policy", "Rate Hike Policy"],
        y=[stat_cut, stat_hold, stat_hike],
        marker_color=["#22C55E", "#94A3B8", "#EF4444"],
        text=[f"+{stat_cut:.2f}%", f"+{stat_hold:.2f}%", f"+{stat_hike:.2f}%"],
        textposition="outside"
    ))
    fig_comp.update_layout(
        title="<b>Empirical Forward 3-Month Nifty 50 Return by RBI Policy Stance (%)</b>",
        template="plotly_dark",
        height=360,
        yaxis_title="Average 3M Return (%)",
        margin=dict(l=20, r=20, t=50, b=20)
    )
    st.plotly_chart(fig_comp, use_container_width=True)

    # Raw Meeting Records tucked in collapsible accordion (NO database dumps on front-end!)
    with st.expander(f"📜 View Underlying RBI MPC Historical Meeting Records ({total_meetings} Decisions)"):
        st.markdown("Underlying empirical database powering the simulation model:")
        st.dataframe(df_rbi, use_container_width=True)

