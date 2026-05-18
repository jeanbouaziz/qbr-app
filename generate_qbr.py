@app.route("/generate_qbr", methods=["POST"])
def generate_qbr():

    print("🔥 GENERATE_QBR HIT")

    import pandas as pd
    import matplotlib.pyplot as plt
    from pptx import Presentation
    from pptx.util import Inches
    from openpyxl import load_workbook
    import anthropic

    # -----------------------
    # Get uploaded files
    # -----------------------
    client_name = request.form.get("client")

    excel_file = request.files["excel"]
    ppt_file = request.files["ppt"]

    print("Client:", client_name)

    # -----------------------
    # Save uploaded files
    # -----------------------
    excel_path = "temp.xlsx"
    ppt_path = "temp_template.pptx"

    excel_file.save(excel_path)
    ppt_file.save(ppt_path)

    print("✅ Files saved")

    # -----------------------
    # Read Excel - Sheet 1: Raw data
    # -----------------------
    wb = load_workbook(excel_path)

    df_raw = pd.read_excel(excel_path, sheet_name="PASTE Company data")
    df_raw.columns = df_raw.columns.str.strip()

    df_auto = pd.read_excel(excel_path, sheet_name="AUTO Intermediary table")
    df_auto.columns = df_auto.columns.str.strip()

    print("✅ Excel loaded")

    # -----------------------
    # Calculate averages from raw data sheet
    # -----------------------
    avg_collected       = df_raw["Collected posts"].mean()
    avg_approved        = df_raw["Total Approved Posts"].mean()
    avg_engagement_rate = df_raw["Engagement Rate (%)"].mean()
    avg_order_share     = df_raw["Order share (%)"].mean()
    avg_order_value_uplift  = df_raw["Order value uplift (%)"].mean()
    avg_order_size_uplift   = df_raw["Order size uplift (%)"].mean()

    # Rights requests (comments + DMs combined)
    avg_rights_requests = (
        df_raw["Rights Request Sent By Comment"] +
        df_raw["Rights Request Sent By DM"]
    ).mean()

    # Post engagement ratio (post engagements / total engagements)
    avg_post_engagement = (
        df_raw["Post Engagements"] / df_raw["Engagements"] * 100
    ).mean()

    # CTA engagement ratio (CTA / total engagements)
    avg_cta_engagement = (
        df_raw["CTA Engagement"] / df_raw["Engagements"] * 100
    ).mean()

    # Approval rate
    avg_approval_rate = (avg_approved / avg_collected * 100) if avg_collected else 0

    # Average times each approved post is distributed
    avg_distributed_per_post = (
        df_raw["Distributed Post To A Flow"] / df_raw["Total Approved Posts"]
    ).mean()

    # Average posts distributed by adding products
    avg_distributed_products = df_raw["Added Product to A Post"].mean()

    print("✅ Raw data averages calculated")

    # -----------------------
    # Calculate averages from AUTO sheet (cost metrics)
    # -----------------------
    avg_cost_per_1k         = df_auto["Cost per 1k impressions"].mean()
    avg_cost_per_engagement = df_auto["Cost per post engagement"].mean()
    avg_cost_per_click      = df_auto["Cost per product click"].mean()
    avg_3pct_attribution    = df_auto["3% attribution"].mean()
    avg_5pct_attribution    = df_auto["5% attribution"].mean()
    avg_7pct_attribution    = df_auto["7% attribution"].mean()

    print("✅ Cost averages calculated")

    # -----------------------
    # Call Claude API
    # -----------------------
    print("⏳ Calling Claude API...")

    anthropic_client = anthropic.Anthropic()

    kpi_summary = f"""
    All figures are averages over the last 12 months:
    - Collected posts/month: {avg_collected:.0f}
    - Approved posts/month: {avg_approved:.0f}
    - Approval rate: {avg_approval_rate:.1f}%
    - Avg times each post is distributed: {avg_distributed_per_post:.1f}
    - Posts distributed via product tagging: {avg_distributed_products:.0f}
    - Rights requests sent/month: {avg_rights_requests:.1f}
    - Engagement rate: {avg_engagement_rate:.2f}%
    - Post engagement ratio: {avg_post_engagement:.1f}%
    - CTA engagement ratio: {avg_cta_engagement:.1f}%
    - Cost per 1k impressions: €{avg_cost_per_1k:.2f}
    - Cost per post engagement: €{avg_cost_per_engagement:.2f}
    - Cost per click: €{avg_cost_per_click:.2f}
    - ROI - 3% attribution: {avg_3pct_attribution:.2f}x
    - ROI - 5% attribution: {avg_5pct_attribution:.2f}x
    - ROI - 7% attribution: {avg_7pct_attribution:.2f}x
    - Order value uplift: {avg_order_value_uplift:.1f}%
    - Order size uplift: {avg_order_size_uplift:.1f}%
    - Order share: {avg_order_share:.2f}%
    """

    prompt = f"""
    You are writing content for a Quarterly Business Review (QBR) presentation for
    {client_name}, a Flowbox client. Flowbox is a UGC (user-generated content) and social commerce platform.

    Based on the following KPIs, generate three things:

    1. NARRATIVE: A 2-3 sentence executive summary of overall performance.
       Reference {client_name} by name. Be specific with numbers. Tone: professional but warm.

    2. INSIGHTS: 3 bullet points of the most interesting observations from the data.
       Focus on standout metrics — good or bad.

    3. TALKING_POINTS: 3 bullet points a CSM can use when presenting to {client_name}.
       Mix value delivered with concrete next step recommendations.

    KPI Data:
    {kpi_summary}

    Respond ONLY in this exact format (no extra text):
    NARRATIVE: <your narrative here>
    INSIGHTS:
    - <insight 1>
    - <insight 2>
    - <insight 3>
    TALKING_POINTS:
    - <point 1>
    - <point 2>
    - <point 3>
    """

    message = anthropic_client.messages.create(
        model="claude-opus-4-5",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}]
    )

    ai_response = message.content[0].text
    print("✅ Claude responded")
    print(ai_response)

    # -----------------------
    # Parse AI response
    # -----------------------
    ai_narrative = ""
    ai_insights = ""
    ai_talking_points = ""

    lines = ai_response.strip().split("\n")
    current_section = None

    for line in lines:
        if line.startswith("NARRATIVE:"):
            ai_narrative = line.replace("NARRATIVE:", "").strip()
            current_section = "narrative"
        elif line.startswith("INSIGHTS:"):
            current_section = "insights"
        elif line.startswith("TALKING_POINTS:"):
            current_section = "talking_points"
        elif line.startswith("- ") and current_section == "insights":
            ai_insights += line + "\n"
        elif line.startswith("- ") and current_section == "talking_points":
            ai_talking_points += line + "\n"

    print("✅ AI content parsed")

    # -----------------------
    # Load PPT template
    # -----------------------
    ppt = Presentation(ppt_path)

    # -----------------------
    # All placeholder replacements
    # -----------------------
    replacements = {
        "{{COLLECTED}}":                        f"{avg_collected:.0f}",
        "{{APPROVED}}":                         f"{avg_approved:.0f}",
        "{{APPROVALRATE}}":                     f"{avg_approval_rate:.1f}%",
        "{{AVERAGEDISTRIBUTED}}":               f"{avg_distributed_per_post:.1f}",
        "{{DISTRIBUTED}}":                      f"{avg_distributed_products:.0f}",
        "{{RIGHTREQUESTS}}":                    f"{avg_rights_requests:.1f}",
        "{{AVERAGEENGAGEMENTRATE}}":            f"{avg_engagement_rate:.2f}",
        "{{POSTENGAGEMENT}}":                   f"{avg_post_engagement:.1f}",
        "{{CTAENGAGEMENT}}":                    f"{avg_cta_engagement:.1f}",
        "{{PER1KIMPRESSIONS}}":                 f"{avg_cost_per_1k:.2f}",
        "{{PERPOSTENGAGEMEN}}":                 f"{avg_cost_per_engagement:.2f}",
        "{{COSTPERCLICK}}":                     f"€{avg_cost_per_click:.2f}",
        "{{AVERAGEPERPOSTENGAGEMENT}}":         f"€{avg_cost_per_engagement:.2f}",
        "{{AVERAGECOSTPERCLICK}}":              f"{avg_cost_per_click:.2f}",
        "{{3%ATTRIBUTIONOFASSISTEDSALES}}":     f"{avg_3pct_attribution:.2f}",
        "{{5%ATTRIBUTIONOFASSISTEDSALES}}":     f"{avg_5pct_attribution:.2f}",
        "{{10%ATTRIBUTIONOFASSISTEDSALES}}":    f"{avg_7pct_attribution:.2f}",
        "{{ORDERVALUEUPLIFTAVERAGE}}":          f"{avg_order_value_uplift:.1f}",
        "{{ORDERSIZEVALUEUPLIFT}}":             f"{avg_order_size_uplift:.1f}",
        "{{ORDERSHAREAVERAGE}}":                f"{avg_order_share:.2f}",
        "{{AI_SUMMARY}}":                       ai_narrative,
        "{{AI_INSIGHTS}}":                      ai_insights.strip(),
        "{{AI_TALKING_POINTS}}":                ai_talking_points.strip(),
    }

    # -----------------------
    # Replace placeholders in PPT
    # -----------------------
    for slide in ppt.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in replacements.items():
                            if key in run.text:
                                run.text = run.text.replace(key, value)

    print("✅ Text replaced")

    # -----------------------
    # Save and return PPT
    # -----------------------
    output_path = "final_qbr.pptx"
    ppt.save(output_path)

    print("✅ PPT GENERATED SUCCESSFULLY")

    return send_file(output_path, as_attachment=True)
