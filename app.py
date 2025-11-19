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

# Utility function to write plain text output (used for Data Model Audit)
def write_text_file(text):
    output = io.BytesIO()
    output.write(text.encode("utf-8"))
    output.seek(0)
    return output

# Governance Lineage Logic
def generate_governance_lineage(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_rules = pd.read_excel(xls, sheet_name="BUSINESS RULES", engine="openpyxl")
    df_conditions = pd.read_excel(xls, sheet_name="BUSINESS CONDITIONS", engine="openpyxl")
    df_mapping = pd.read_excel(xls, sheet_name="GOVERNANCE MAPPING", engine="openpyxl")

    # --- FIX: Handle duplicate "CONTEXT TYPE || CONTEXT NAME" columns ---
    df_contexts_raw = pd.read_excel(xls, sheet_name="CONTEXTS", header=0, engine="openpyxl")

    # Ensure duplicate column names get unique suffixes (.1, .2, etc.)
    cols = list(df_contexts_raw.columns)
    new_cols = []
    seen = {}
    for col in cols:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}.{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    df_contexts_raw.columns = new_cols

    # Select the SECOND occurrence (Column D)
    context_type_name_col = [col for col in df_contexts_raw.columns if col.startswith("CONTEXT TYPE || CONTEXT NAME")][1]

    df_contexts = df_contexts_raw.rename(columns={
        "NAME": "CONTEXT NAME",
        context_type_name_col: "CONTEXT TYPE AND NAME"
    })
    df_contexts = df_contexts[[
        "CONTEXT NAME", "CONTEXT TYPE AND NAME",
        "WORKFLOW ACTIVITY", "WORKFLOW ACTIVITY ACTION(s)", "WORKFLOW ACTIVITY CRITERIA"
    ]]
    # --- END FIX ---

    df_rules = df_rules.rename(columns={"NAME": "RULE NAME", "DISPLAY NAME": "RULE DISPLAY NAME"})
    df_rules = df_rules[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME", "IS ENABLED?"]]
    df_rules = df_rules[df_rules["IS ENABLED?"] == "Yes"]

    df_conditions = df_conditions.rename(columns={"NAME": "CONDITION NAME", "DISPLAY NAME": "CONDITION DISPLAY NAME"})
    df_conditions = df_conditions[[
        "CONDITION NAME", "MAPPED BUSINESS RULE(s)", "IMPACTED ROLES", "IMPACTED ATTRIBUTES",
        "IMPACTED RELATIONSHIPS", "CONDITION DISPLAY NAME", "IS ENABLED?"
    ]]
    df_conditions = df_conditions[df_conditions["IS ENABLED?"] == "Yes"]

    df_mapping["FOR CONTEXT"] = df_mapping["FOR CONTEXT"].astype(str)
    df_mapping = df_mapping[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT", "IS ENABLED?"]]
    df_mapping = df_mapping[df_mapping["IS ENABLED?"] == "Yes"]

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


# Keyword Analysis Logic (simplified)
def generate_keyword_analysis(file):
    keywords = ["GetEntityBusinessConditionScore", "ContextsHave", "IsContextChanged", "GetContextPath", "GetEntityId", "GetEntityIds", "GetEntityType", "GetEntityVersion", "GetEntityName", "GetEntityProperty", "GetCurrentWorkflowAssignedUser", "GetLocalesOfChangedData", "GetCurrentWorkflowStep", "GetAllContexts", "GetChildContexts", "HasContextLinks", "HaveErrorsInContext", "IsEntityDeleted", "IsEntityInWorkflow", "IsEntityInWorkflowInContext", "SetEntityProperty", "GetConfigKeyValue", "GenerateUniqueId", "GetWeekOfYear", "AttributesHaveErrorsInContext", "GetAttributeValue", "GetAttributeValueFromContext", "GetAttributeValuesFromContext", "GetAttributeValueWithDefault", "GetAttributeValues", "GetAttributeValuesWithDefault", "GetAttributeValuesWithDefaultFromContext", "GetAttributeValueWithDefaultFromContext", "GetAttributeValueReferenceId", "GetAttributeValueReferenceIds", "GetAttributeValueProperty", "GetEntityAttributeValuesById", "GetEntityAttributeValueById", "GetEntityAttributeValueByIdInContext", "GetNestedAttributeComputedValue", "GetNestedAttributeValues", "GetNestedAttributeRow", "GetNestedAttributeRows", "DeleteNestedAttributeRows", "GetEntityNestedAttributeRow", "GetEntityNestedAttributeRows", "GetNestedAttributeValueReferenceID", "GetChildAttributefromNestedRowString", "GetRelatedEntityIdByAttributeValue", "GetRelatedEntityIdByAttributeValueFromContext", "GetRelatedEntityIdsByAttributeValue", "GetRelatedEntityIdsByAttributeValueFromContext", "HaveAnyAttributesChanged", "HaveAnyAttributesChangedInContext", "HaveAnyRelationshipAttributesChanged", "HaveAttributesChanged", "HaveAttributesChangedInContext", "IsAttributeLocalizable", "ValidateEmptyAttributes", "ValidateEmptyAttributesInContext", "DeleteAttribute", "DeleteEntityAttribute", "DeleteAttributeInContext", "DeleteRelationshipAttributeInContext", "GetExternalSourceOfAttribute", "GetEntitiesAttributesValues", "HaveOnlySpecifiedAttributesChanged", "GetMappedAttributeNames", "GetDeeplyNestedAttributeJSON", "GetEntityDeeplyNestedAttributeJSON", "GetAttributePreviousValues", "AreRelationshipsDeleted", "CheckIfAllRelationshipAttributeValueIs", "CheckIfAnyRelationshipAttributeValueIs", "CheckIfAllRelatedEntityAttributeValueIs", "CheckIfAnyRelatedEntityAttributeValueIs", "GetCurrentRelatedEntityIds", "GetRelatedEntityIdForContext", "GetRelatedEntityId", "GetRelatedEntityIds", "GetRelatedEntityIdsForContext", "GetRelatedEntityIdByRelationshipAttributeValue", "GetRelatedEntityIdsByRelationshipAttributeValue", "GetRelatedEntityIdByRelationshipAttributeValueFromContext", "GetRelatedEntityIdsByRelationshipAttributeValueFromContext", "GetRelationshipAttributevalue", "GetRelationshipAttributevalues", "HaveRelationships", "HaveRelationshipsInContext", "HaveRelationshipsChanged", "RelationshipsHaveErrorsInContext", "RelationshipsCountInContext", "ValidateEmptyRelationshipAttributes", "ValidateEmptyRelationshipAttributesInContext", "ValidateEmptyAttributesForRelatedEntities", "ValidateEmptyAttributesForRelatedEntitiesInContext", "WhereUsedRelationship", "IsInheritanceBlocked", "GetWhereUsedEntityIds", "IsCurrentUserInRole", "CurrentUser", "GetImpersonateUser", "GetUserOwnershipData", "GetUserOwnershipEditData", "GetUserOwnershipDataCollection", "GetUserOwnershipEditDataCollection", "GetUserProperty", "StopBRExecution", "ValidateExternalLink", "GetClientAttributesFromRequest", "GetDefaultLocaleForTenant", "GetGlobalVariable", "GetRestAPIResponse", "GetUniqueId", "JoinStringCollection", "SetVariable", "SetGlobalVariable", "ValidateByRegex", "GetOriginatingClientId", "GetClientId", "ValidateGTINCheckDigit", "ValidateISBNCheckDigit", "CalculateGTINCheckDigit", "GetValueByJsonPath", "ExtractUOMInfo", "ValidateLuhnAlgorithm", "HasSrcAloneChanged", "URLEncode", "AddToContext", "DeleteContext", "AddNestedAttributeRow", "AddNestedAttributeRowInContext", "SetAttributeValue", "SetAttributeValueInContext", "SetAttributeValues", "SetAttributeValuesInContext", "SetNestedChildAttributeByCondition", "SetDeeplyNestedAttributeJSON", "DeleteRelationships", "SetRelationshipAttribute", "SetRelationshipAttributeFromRelatedEntity", "AddRelationshipInContextByEntityId", "CopyAttributeValueToGovern", "GetBusinessConditionStatus", "GetEntityBusinessConditionStatus", "AddAttributeError", "AddAttributeInformation", "AddContextError", "AddContextInformation", "AddContextWarning", "AddAttributeErrorInContext", "AddAttributeInformationInContext", "AddAttributeWarningInContext", "AddRelationshipAttributeError", "AddRelationshipAttributeInformation", "AddRelationshipAttributeWarning", "AddRelationshipAttributeErrorInContext", "AddRelationshipAttributeInformationInContext", "AddRelationshipAttributeWarning", "AddRelationshipError", "AddRelationshipInformation", "AddRelationshipWarning", "AddRelationshipInformationInContext", "AddRelationshipErrorInContext", "AddRelationshipWarningInContext", "AddAttributeWarning", "ValidatePhone", "ChangeAssignment", "ChangeAssignmentInContext", "InitiateExport", "InitiateExportInContext", "InitiateExportInLocale", "InitiateExportInContextAndLocale", "InitiateExportForDeletedEntity", "InitiateExportForDeletedEntityInContext", "InitiateExportForEntity", "InitiateExportForRelatedEntity", "InitiateExportForDeletedEntityInContextAndLocale", "InvokeWorkflow", "InvokeWorkflowInContext", "ResumeWorkflow", "ResumeWorkflowInContext", "ScheduleEntityForExport", "ScheduleEntityForGraphProcessing", "ScheduleWhereUsedEntitiesForGraphProcessing", "SendEntityForGraphProcessing", "SendWhereUsedEntitiesForGraphProcessing", "SendEmail", "CreateSnapshot", "RestoreSnapshot", "ExportApprovedVersion", "CreateAndExportApprovedVersion", "CreateEntity", "DeleteEntity", "GetBusinessConditionStatus", "GetEntityBusinessConditionStatus", "ManageAddress", "GetWorkflowComment", "GetEntityCurrentWorkflowStep", "EndWorkflow", "GenerateVariants", "ScheduleOrSendEntityForGraphProcessing", "SetEntityAttributeValue", "SetEntityAttributeValueForContext", "AddEntityNestedAttributeRow", "SetEntityDeeplyNestedAttributeJSON", "CheckIfAnyWhereUsedEntityAttributeValueIs", "GetChangedNestedAttributeRows", "GetDeletedNestedAttributeRows", "ResumeRelatedEntityWorkflow", "ScheduleRelatedEntitiesForGraphProcessing", "SendRelatedEntitiesForGraphProcessing", "SetRelatedEntityAttributeValue", "SetRelatedEntityAttributeValueForContext", "WhereUsedRelationshipsCountInContext", "GetConnectorState", "SetConnectorState", "InvokeConnectorState", "AttributeInContext", "SortedAttributeValues", "SortedAttributeValuesFromContext", "URLEncode", "GetApplicationURL", "CurrentWorkflowStepStartDate", "ContextType", "ContextPath"
        # ... include all remaining keywords from your file
        "URLEncode", "GetApplicationURL", "CurrentWorkflowStepStartDate", "ContextType", "ContextPath"
    ]
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_rules = pd.read_excel(xls, sheet_name="BUSINESS RULES", engine="openpyxl")

    # Filter enabled rules
    df_rules = df_rules[df_rules["IS ENABLED?"] == "Yes"]
    df_rules["DEFINITION"] = df_rules["DEFINITION"].astype(str).str.strip()


results = []
for keyword in keywords:
    count = df_rules[df_rules["DEFINITION"].str.contains(keyword, case=False, na=False)]["NAME"].nunique()
    if count > 0:  # ✅ Only include keywords that are used
        results.append({"Keyword": keyword, "Count of Matching Rules": count})

# Create DataFrame only with used keywords
df_output = pd.DataFrame(results)
if not df_output.empty:
    df_output = df_output.sort_values(by="Count of Matching Rules", ascending=False)
else:
    df_output = pd.DataFrame([{"Keyword": "No matches found", "Count of Matching Rules": 0}])


    return write_clean_excel(df_output)


# Unused Business Rules Logic
def generate_unused_business_rules(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_rules = pd.read_excel(xls, sheet_name="BUSINESS RULES", engine="openpyxl")
    df_conditions = pd.read_excel(xls, sheet_name="BUSINESS CONDITIONS", engine="openpyxl")
    df_mapping = pd.read_excel(xls, sheet_name="GOVERNANCE MAPPING", engine="openpyxl")

    df_rules = df_rules[df_rules["IS ENABLED?"] == "Yes"]
    rule_names = set(df_rules["NAME"].dropna().astype(str))

    condition_rules = df_conditions[df_conditions["IS ENABLED?"] == "Yes"]["MAPPED BUSINESS RULE(s)"].dropna().astype(str)
    mapped_condition_rules = set()
    for entry in condition_rules:
        mapped_condition_rules.update([rule.strip() for rule in entry.split("||")])

    mapping_rules = df_mapping[df_mapping["IS ENABLED?"] == "Yes"]["MAPPED BUSINESS RULE"].dropna().astype(str)
    mapped_mapping_rules = set()
    for entry in mapping_rules:
        mapped_mapping_rules.update([rule.strip() for rule in entry.split("||")])

    all_mapped_rules = mapped_condition_rules.union(mapped_mapping_rules)
    unused_rules = sorted(list(rule_names - all_mapped_rules))
    df_unused = pd.DataFrame({"Unused Business Rules": unused_rules})
    return write_clean_excel(df_unused)

# -------------------------------
# Data Model Lineage Logic
# -------------------------------
def generate_data_model_lineage(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_attr = pd.read_excel(xls, sheet_name="ATTRIBUTES", engine="openpyxl")
    df_ear = pd.read_excel(xls, sheet_name="E-A-R MODEL", engine="openpyxl")

    df_attr = df_attr[["NAME", "DISPLAY NAME", "DATA TYPE", "USES REFERENCE DATA", "PATH ROOT NODE"]]
    df_ear = df_ear[["MAPPED ATTRIBUTE", "ENTITY"]]
    df_ear = df_ear[df_ear["ENTITY"].notna() & (df_ear["ENTITY"].astype(str).str.strip() != "")]

    df_merged = df_attr.merge(df_ear, left_on="NAME", right_on="MAPPED ATTRIBUTE", how="inner")
    df_merged = df_merged[["NAME", "ENTITY", "DISPLAY NAME", "DATA TYPE", "USES REFERENCE DATA", "PATH ROOT NODE"]]
    df_merged = df_merged.sort_values(by="NAME")

    # Return Excel file as BytesIO
    return write_clean_excel(df_merged)


# -------------------------------
# Data Model Audit Report Logic
# -------------------------------
def generate_data_model_audit(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_entities = pd.read_excel(xls, sheet_name="ENTITIES", engine="openpyxl")
    df_relationships = pd.read_excel(xls, sheet_name="RELATIONSHIPS", engine="openpyxl")
    df_attributes = pd.read_excel(xls, sheet_name="ATTRIBUTES", engine="openpyxl")
    df_ear = pd.read_excel(xls, sheet_name="E-A-R MODEL", engine="openpyxl")

    # Unused Entities
    entity_names = set(df_entities["NAME"].dropna().astype(str))
    used_entities = set(df_ear["ENTITY"].dropna().astype(str))
    unused_entities = sorted(list(entity_names - used_entities))

    # Unused Relationships
    rel_names = set(df_relationships["NAME"].dropna().astype(str))
    used_rels = set(df_ear["MAPPED RELATIONSHIP"].dropna().astype(str))
    unused_relationships = sorted(list(rel_names - used_rels))

    # Unmapped Attributes
    df_nested_parents = df_attributes[df_attributes["DISPLAY TYPE"] == "nestedgrid"]
    nested_parent_names = set(df_nested_parents["NAME"].dropna().astype(str))
    df_unmapped = df_attributes[~df_attributes["GROUP"].isin(nested_parent_names)]
    attr_names = set(df_unmapped["NAME"].dropna().astype(str))
    mapped_attrs = set(df_ear["MAPPED ATTRIBUTE"].dropna().astype(str))
    unmapped_attributes = sorted(list(attr_names - mapped_attrs))

    # Nestedgrid Attributes Without Identifier
    missing_identifiers = []
    for parent in nested_parent_names:
        children = df_attributes[df_attributes["GROUP"] == parent]
        identifiers = children["IS NESTED GROUP IDENTIFIER?"].dropna().astype(str).str.lower()
        if not any(identifiers == "yes"):
            missing_identifiers.append(parent)
    missing_identifiers = sorted(missing_identifiers)

    # Format output
    report = "Unused Entities:\n" + "\n".join(f"- {e}" for e in unused_entities) + "\n\n"
    report += "Unused Relationships:\n" + "\n".join(f"- {r}" for r in unused_relationships) + "\n\n"
    report += "Unmapped Attributes:\n" + "\n".join(f"- {a}" for a in unmapped_attributes) + "\n\n"
    report += "Nestedgrid Attributes Without Identifier:\n" + "\n".join(f"- {n}" for n in missing_identifiers)

    # Return text file as StringIO
    return write_text_file(report)


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

        unused_output = generate_unused_business_rules(gov_file)
        st.download_button("Unused Business Rules in Governance Model", data=unused_output, file_name="unused_business_rules.xlsx")


with col2:
    st.header("Upload Dynamic Authorization Model")
    st.markdown("Generate data lineage document from your Dynamic Authorization model Excel file.")
    auth_file = st.file_uploader("Upload Authorization Excel (.xlsm)", key="auth")
    if auth_file:
        output = generate_auth_lineage(auth_file)
        st.download_button("Download Authorization Lineage", data=output, file_name="dynamic_auth_lineage_output.xlsx")

# -------------------------------
# New Section: Data Model
# -------------------------------

col3, _ = st.columns([1, 1])  # keeps same centered width as other uploaders
with col3:
    st.header("Upload Data Model")
    st.markdown("Generate data lineage document and Audit report from your Data model Excel file.")
     # 🔹 Brief description of Audit Report contents
    st.markdown("""
    <div style='background-color:#f8f9fa; border-left:4px solid #4a90e2; padding:10px; border-radius:8px;'>
        <b>Audit Report includes:</b><br>
        • Unused Entities<br>
        • Unused Relationships<br>
        • Unmapped Attributes<br>
        • Nestedgrid Attributes Without Identifier
    </div>
    """, unsafe_allow_html=True)
    data_file = st.file_uploader("Upload Data Model Excel (.xlsx)", key="data_model")

    if data_file:
        lineage_output = generate_data_model_lineage(data_file)
        audit_output = generate_data_model_audit(data_file)

        st.download_button(
            "Download Data Model Lineage Document",
            data=lineage_output,
            file_name="Datamodel_lineage_document.xlsx"
        )

        st.download_button(
            "Download Data Model Audit Report",
            data=audit_output,
            file_name="Datamodel_audit_report.txt"
        )
