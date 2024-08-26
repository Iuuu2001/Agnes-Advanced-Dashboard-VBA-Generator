import streamlit as st
import pandas as pd
import textwrap

st.set_page_config(layout="wide", initial_sidebar_state="collapsed")

st.title("VBA Generator for advanced dashboard")

st.write("This is a VBA generator for advanced dashboards")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:

    mainPage  = pd.read_excel(uploaded_file, sheet_name="主頁",header=0)
    subPage   = pd.read_excel(uploaded_file, sheet_name="分頁",header=0)
    innerPage = pd.read_excel(uploaded_file, sheet_name="內頁",header=0)

    st.write("Excel file has been uploaded successfully.")

    no_of_mainPage = len(mainPage)
    no_of_subPage  = len(subPage)
    no_of_innerPage= len(innerPage)

    VBA = []

    # Handle mainpage to subpage
    if no_of_mainPage > 0 and no_of_subPage > 0:
        for i in range(len(subPage)):
            code_name = subPage["Code名稱"][i]
            page_name = subPage["分頁名稱"][i]

            VBA.append(textwrap.dedent(f"""
            Sub main_to_{code_name}()
                ThisWorkbook.Sheets("{page_name}").Visible = True
                ThisWorkbook.Sheets("{page_name}").Activate
                ActiveSheet.Range("A1").Select
            End Sub
            """))

    # Handle subpage to subpage and subpage to mainpage
    if no_of_subPage > 0:
        for i in range(len(subPage)):
            code_name1 = subPage["Code名稱"][i]
            page_name1 = subPage["分頁名稱"][i]

            for j in range(len(subPage)):
                code_name2 = subPage["Code名稱"][j]
                page_name2 = subPage["分頁名稱"][j]

                if i != j:
                    VBA.append(textwrap.dedent(f"""
                    Sub {code_name1}_{code_name2}()
                        ThisWorkbook.Sheets("{page_name1}").Visible = xlVeryHidden
                        ThisWorkbook.Sheets("{page_name2}").Activate
                        ActiveSheet.Range("A1").Select
                    End Sub
                    """))

            if no_of_mainPage > 0:
                main_page_name = mainPage["主頁名稱"][0]

                VBA.append(textwrap.dedent(f"""
                Sub {code_name1}_main()
                    ThisWorkbook.Sheets("{page_name1}").Visible = xlVeryHidden
                    ThisWorkbook.Sheets("{main_page_name}").Activate
                    ActiveSheet.Range("A1").Select
                End Sub
                """))

    # Handle innerPage to corresponding subpage / innerPage to subpage and innerPage to mainpage
    if no_of_innerPage > 0:
        for i in range(len(innerPage)):
            innerPage_name = innerPage["內頁名稱"][i]
            innerPage_codeName = innerPage["Code名稱"][i]
            innerPage_cor = innerPage["所屬分頁名稱"][i]
            innerPage_corCodeName = innerPage["所屬分頁code名稱"][i]
            innerPage_connect = innerPage["連接分頁(Y/N)"][i]

            if innerPage_connect == "Y":
                for j in range(len(subPage)):
                    code_name = subPage["Code名稱"][j]
                    page_name = subPage["分頁名稱"][j]

                    VBA.append(textwrap.dedent(f"""
                    Sub INNER_{innerPage_codeName}_{code_name}()
                        ThisWorkbook.Sheets("{innerPage_name}").Visible = xlVeryHidden
                        ThisWorkbook.Sheets("{page_name}").Activate
                        ActiveSheet.Range("A1").Select
                    End Sub
                    """))

                if no_of_mainPage > 0:
                    main_page_name = mainPage["主頁名稱"][0]

                    VBA.append(textwrap.dedent(f"""
                    Sub INNER_{innerPage_codeName}_main()
                        ThisWorkbook.Sheets("{innerPage_name}").Visible = xlVeryHidden
                        ThisWorkbook.Sheets("{main_page_name}").Activate
                        ActiveSheet.Range("A1").Select
                    End Sub
                    """))
            else:
                VBA.append(textwrap.dedent(f"""
                Sub INNER_{innerPage_codeName}_{innerPage_corCodeName}()
                    ThisWorkbook.Sheets("{innerPage_name}").Visible = xlVeryHidden
                    ThisWorkbook.Sheets("{innerPage_cor}").Activate
                    ActiveSheet.Range("A1").Select
                End Sub
                """))

    # Display generated VBA code
    st.subheader("Generated VBA")
    st.code("\n".join(VBA), language='visualBasic (visual-basic)')