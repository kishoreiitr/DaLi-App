import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Lineage Generator", layout="wide")

# Utility function to write Excel
def write_clean_excel(df):
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    ws.auto_filter.ref = ws.dimensions
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Governance Lineage Logic
def generate_governance_lineage(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_rules = pd.read_excel(xls, sheet_name="BUSINESS RULES", engine="openpyxl")
    df_conditions = pd.read_excel(xls, sheet_name="BUSINESS CONDITIONS", engine="openpyxl")
    df_mapping = pd.read_excel(xls, sheet_name="GOVERNANCE MAPPING", engine="openpyxl")
    df_contexts = pd.read_excel(xls, sheet_name="CONTEXTS", engine="openpyxl")

    df_rules = df_rules.rename(columns={"NAME": "RULE NAME", "DISPLAY NAME": "RULE DISPLAY NAME"})
    df_rules = df_rules[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME", "IS ENABLED?"]]
    df_rules = df_rules[df_rules["IS ENABLED?"] == "Yes"]

    df_conditions = df_conditions.rename(columns={"NAME": "CONDITION NAME", "DISPLAY NAME": "CONDITION DISPLAY NAME"})
    df_conditions = df_conditions[["CONDITION NAME", "MAPPED BUSINESS RULE(s)", "IMPACTED ROLES", "IMPACTED ATTRIBUTES",
                                   "IMPACTED RELATIONSHIPS", "CONDITION DISPLAY NAME", "IS ENABLED?"]]
    df_conditions = df_conditions[df_conditions["IS ENABLED?"] == "Yes"]

    df_mapping["FOR CONTEXT"] = df_mapping["FOR CONTEXT"].astype(str)
    df_mapping = df_mapping[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT", "IS ENABLED?"]]
    df_mapping = df_mapping[df_mapping["IS ENABLED?"] == "Yes"]

    df_contexts = df_contexts.rename(columns={"NAME": "CONTEXT NAME", "CONTEXT TYPE || CONTEXT NAME": "CONTEXT TYPE AND NAME"})
    df_contexts = df_contexts[["CONTEXT NAME", "CONTEXT TYPE AND NAME", "WORKFLOW ACTIVITY",
                               "WORKFLOW ACTIVITY ACTION(s)", "WORKFLOW ACTIVITY CRITERIA"]]

    lineage_records = []

    for _, rule in df_rules.iterrows():
        rule_name = rule["RULE NAME"]
        matched_conditions = df_conditions[df_conditions["MAPPED BUSINESS RULE(s)"].str.contains(rule_name, na=False)]

        matched = False
        if not matched_conditions.empty:
            for _, cond in matched_conditions.iterrows():
                cond_name = cond["CONDITION NAME"]
                context_keys = [f"{cond_name}{'' if i == 0 else i}Context" for i in range(16)]
                matched_mappings = df_mapping[df_mapping["FOR CONTEXT"].isin(context_keys)]
                for _, map_row in matched_mappings.iterrows():
                    context_row = df_contexts[df_contexts["CONTEXT NAME"] == map_row["FOR CONTEXT"]]
                    context_data = context_row.iloc[0].to_dict() if not context_row.empty else {
                        "CONTEXT NAME": "", "CONTEXT TYPE AND NAME": "", "WORKFLOW ACTIVITY": "",
                        "WORKFLOW ACTIVITY ACTION(s)": "", "WORKFLOW ACTIVITY CRITERIA": ""
                    }
                    lineage_records.append({
                        "RULE NAME": rule["RULE NAME"], "TYPE": rule["TYPE"], "DEFINITION": rule["DEFINITION"],
                        "RULE DISPLAY NAME": rule["RULE DISPLAY NAME"],
                        "CONDITION NAME": cond["CONDITION NAME"], "IMPACTED ROLES": cond["IMPACTED ROLES"],
                        "IMPACTED ATTRIBUTES": cond["IMPACTED ATTRIBUTES"], "IMPACTED RELATIONSHIPS": cond["IMPACTED RELATIONSHIPS"],
                        "CONDITION DISPLAY NAME": cond["CONDITION DISPLAY NAME"],
                        "ENTITY": map_row["ENTITY"], "MAPPED BUSINESS RULE": map_row["MAPPED BUSINESS RULE"],
                        "MAPPED BUSINESS CONDITION": map_row["MAPPED BUSINESS CONDITION"], "FOR CONTEXT": map_row["FOR CONTEXT"],
                        **context_data
                    })
                    matched = True
        if not matched:
            context_keys = [f"{rule_name}{'' if i == 0 else i}Context" for i in range(16)]
            matched_mappings = df_mapping[df_mapping["FOR CONTEXT"].isin(context_keys)]
            if not matched_mappings.empty:
                for _, map_row in matched_mappings.iterrows():
                    context_row = df_contexts[df_contexts["CONTEXT NAME"] == map_row["FOR CONTEXT"]]
                    context_data = context_row.iloc[0].to_dict() if not context_row.empty else {
                        "CONTEXT NAME": "", "CONTEXT TYPE AND NAME": "", "WORKFLOW ACTIVITY": "",
                        "WORKFLOW ACTIVITY ACTION(s)": "", "WORKFLOW ACTIVITY CRITERIA": ""
                    }
                    lineage_records.append({
                        "RULE NAME": rule["RULE NAME"], "TYPE": rule["TYPE"], "DEFINITION": rule["DEFINITION"],
                        "RULE DISPLAY NAME": rule["RULE DISPLAY NAME"],
                        "CONDITION NAME": "", "IMPACTED ROLES": "", "IMPACTED ATTRIBUTES": "", "IMPACTED RELATIONSHIPS": "",
                        "CONDITION DISPLAY NAME": "",
                        "ENTITY": map_row["ENTITY"], "MAPPED BUSINESS RULE": map_row["MAPPED BUSINESS RULE"],
                        "MAPPED BUSINESS CONDITION": map_row["MAPPED BUSINESS CONDITION"], "FOR CONTEXT": map_row["FOR CONTEXT"],
                        **context_data
                    })
                    matched = True
        if not matched:
            fallback_mappings = df_mapping[df_mapping["MAPPED BUSINESS RULE"] == rule_name]
            for _, map_row in fallback_mappings.iterrows():
                lineage_records.append({
                    "RULE NAME": rule["RULE NAME"], "TYPE": rule["TYPE"], "DEFINITION": rule["DEFINITION"],
                    "RULE DISPLAY NAME": rule["RULE DISPLAY NAME"],
                    "CONDITION NAME": "", "IMPACTED ROLES": "", "IMPACTED ATTRIBUTES": "", "IMPACTED RELATIONSHIPS": "",
                    "CONDITION DISPLAY NAME": "",
                    "ENTITY": map_row["ENTITY"], "MAPPED BUSINESS RULE": map_row["MAPPED BUSINESS RULE"],
                    "MAPPED BUSINESS CONDITION": map_row["MAPPED BUSINESS CONDITION"], "FOR CONTEXT": "",
                    "CONTEXT NAME": "", "CONTEXT TYPE AND NAME": "", "WORKFLOW ACTIVITY": "",
                    "WORKFLOW ACTIVITY ACTION(s)": "", "WORKFLOW ACTIVITY CRITERIA": ""
                })

    df_output = pd.DataFrame(lineage_records)
    return write_clean_excel(df_output)

# Dynamic Authorization Logic
def generate_auth_lineage(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_policy = pd.read_excel(xls, sheet_name="POLICY", engine="openpyxl")
    df_mapping = pd.read_excel(xls, sheet_name="POLICY MAPPING", engine="openpyxl")
    df_permissions = pd.read_excel(xls, sheet_name="POLICY PERMISSIONS", engine="openpyxl")

    df_policy = df_policy[df_policy["ENABLED"] == "Yes"]
    df_policy = df_policy[["POLICY", "ENTITY TYPE", "CONDITION"]]

    df_mapping = df_mapping.rename(columns={"POLICY": "MAPPING POLICY", "PERMISSION SET": "MAPPING PERMISSION SET"})
    df_mapping = df_mapping[["MAPPING POLICY", "ROLE", "MAPPING PERMISSION SET"]]

    lineage_records = []

    for _, policy in df_policy.iterrows():
        policy_name = policy["POLICY"]
        mappings = df_mapping[df_mapping["MAPPING POLICY"].str.contains(policy_name, na=False)]
        for _, map_row in mappings.iterrows():
            perm_set = map_row["MAPPING PERMISSION SET"]
            perms = df_permissions[df_permissions["PERMISSION SET"].str.contains(perm_set, na=False)]
            for _, perm_row in perms.iterrows():
                lineage_records.append({
                    "POLICY": policy["POLICY"], "ENTITY TYPE": policy["ENTITY TYPE"], "CONDITION": policy["CONDITION"],
                    "ROLE": map_row["ROLE"],
                    "PERMISSION SET": perm_row["PERMISSION SET"], "ATTRIBUTE": perm_row["ATTRIBUTE"],
                    "RELATIONSHIP": perm_row["RELATIONSHIP"], "PERMISSION": perm_row["PERMISSION"]
                })

    df_output = pd.DataFrame(lineage_records)
    return write_clean_excel(df_output)

# Keyword Analysis Logic with 227 keywords inline
def generate_keyword_analysis(file):
    keywords = [
        "GetEntityBusinessConditionScore", "ContextsHave", "IsContextChanged", "GetContextPath", "GetEntityId",
        "GetEntityIds", "GetEntityType", "GetEntityVersion", "GetEntityName", "GetEntityProperty",
        "GetCurrentWorkflowAssignedUser", "GetLocalesOfChangedData", "GetCurrentWorkflowStep", "GetAllContexts",
        "GetChildContexts", "HasContextLinks", "HaveErrorsInContext", "IsEntityDeleted", "IsEntityInWorkflow",
        "IsEntityInWorkflowInContext", "SetEntityProperty", "GetConfigKeyValue", "GenerateUniqueId", "GetWeekOfYear",
        "AttributesHaveErrorsInContext", "GetAttributeValue", "GetAttributeValueFromContext",
        "GetAttributeValuesFromContext", "GetAttributeValueWithDefault", "GetAttributeValues",
        "GetAttributeValuesWithDefault", "GetAttributeValuesWithDefaultFromContext",
        "GetAttributeValueWithDefaultFromContext", "GetAttributeValueReferenceId", "GetAttributeValueReferenceIds",
        "GetAttributeValueProperty", "GetEntityAttributeValuesById", "GetEntityAttributeValueById",
        "GetEntityAttributeValueByIdInContext", "GetNestedAttributeComputedValue", "GetNestedAttributeValues",
        "GetNestedAttributeRow", "GetNestedAttributeRows", "DeleteNestedAttributeRows", "GetEntityNestedAttributeRow",
        "GetEntityNestedAttributeRows", "GetNestedAttributeValueReferenceID", "GetChildAttributefromNestedRowString",
        "GetRelatedEntityIdByAttributeValue", "GetRelatedEntityIdByAttributeValueFromContext",
        "InitiateExportForDeletedEntityInContextAndLocale", "InvokeWorkflow", "InvokeWorkflowInContext",
        "ResumeWorkflow", "ResumeWorkflowInContext", "ScheduleEntityForExport", "ScheduleEntityForGraphProcessing",
        "ScheduleWhereUsedEntitiesForGraphProcessing", "SendEntityForGraphProcessing",
        "SendWhereUsedEntitiesForGraphProcessing", "SendEmail", "CreateSnapshot", "RestoreSnapshot",
        "ExportApprovedVersion", "CreateAndExportApprovedVersion", "CreateEntity", "DeleteEntity",
        "GetBusinessConditionStatus", "GetEntityBusinessConditionStatus", "ManageAddress", "GetWorkflowComment",
        "GetEntityCurrentWorkflowStep", "EndWorkflow", "GenerateVariants", "ScheduleOrSendEntityForGraphProcessing",
        "SetEntityAttributeValue", "SetEntityAttributeValueForContext", "AddEntityNestedAttributeRow",
        "SetEntityDeeplyNestedAttributeJSON", "CheckIfAnyWhereUsedEntityAttributeValueIs",
        "GetChangedNestedAttributeRows", "GetDeletedNestedAttributeRows", "ResumeRelatedEntityWorkflow",
        "ScheduleRelatedEntitiesForGraphProcessing", "SendRelatedEntitiesForGraphProcessing",
        "SetRelatedEntityAttributeValue", "SetRelatedEntityAttributeValueForContext", "WhereUsedRelationshipsCountInContext",
        "GetConnectorState", "SetConnectorState", "InvokeConnectorState", "AttributeInContext", "SortedAttributeValues",
        "SortedAttributeValuesFromContext", "URLEncode", "GetApplicationURL", "CurrentWorkflowStepStartDate",
        "ContextType", "ContextPath"
    ]

    xls = pd.ExcelFile(file, engine="openpyxl")
    df_rules = pd.read_excel(xls, sheet_name="BUSINESS RULES", engine="openpyxl")
    df_rules = df_rules[df_rules["IS ENABLED?"] == "Yes"]

    results = []
    for keyword in keywords:
        count = df_rules["DEFINITION"].apply(lambda x: keyword in str(x)).sum()
        if count > 0:
            results.append({"Keyword": keyword, "Count of Matching Rules": count})

    df_output = pd.DataFrame(results).sort_values(by="Count of Matching Rules", ascending=False)
    return write_clean_excel(df_output)

# UI Layout
col1, col2 = st.columns(2)

with col1:
    st.header("Upload Governance Model")
    st.markdown("Generate data lineage document from your Governance model Excel file.")
    gov_file = st.file_uploader("Upload Governance Excel (.xlsm)", key="gov")
    if gov_file:
        output = generate_governance_lineage(gov_file)
        st.download_button("Download Governance Lineage", data=output, file_name="Goverance_rules_lineage_output.xlsx")

        keyword_output = generate_keyword_analysis(gov_file)
        st.download_button("Generate Keyword List Document", data=keyword_output, file_name="keywords used in governance model.xlsx")

with col2:
    st.header("Upload Dynamic Authorization Model")
    st.markdown("Generate data lineage document from your Dynamic Authorization model Excel file.")
    auth_file = st.file_uploader("Upload Authorization Excel (.xlsm)", key="auth")
    if auth_file:
        output = generate_auth_lineage(auth_file)
        st.download_button("Download Authorization Lineage", data=output, file_name="dynamic_auth_lineage_output.xlsx")
