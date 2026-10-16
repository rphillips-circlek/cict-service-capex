import sys
from azure.identity import AzureCliCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient

ENDPOINT = "https://circlekdoci.cognitiveservices.azure.com/"
TENANT = "3e41b164-59e6-4ce9-8c15-767e2c81431c"
PDF = r"C:\Users\deepakbalan\Downloads\Cbe Capito - R3645629.pdf"

cred = AzureCliCredential(tenant_id=TENANT)
client = DocumentIntelligenceClient(endpoint=ENDPOINT, credential=cred)
try:
    with open(PDF, "rb") as fh:
        poller = client.begin_analyze_document("prebuilt-invoice", body=fh, content_type="application/pdf")
    result = poller.result()
    print("OK - analyzed. pages:", len(result.pages or []), "| docs:", len(result.documents or []))
    if result.documents:
        flds = result.documents[0].fields or {}
        for key in ("VendorName", "InvoiceId", "InvoiceDate", "PurchaseOrder", "InvoiceTotal", "TotalTax"):
            fv = flds.get(key)
            if fv:
                val = getattr(fv, "content", None) or getattr(fv, "value_string", None)
                print(f"  {key:14}: {val!r}  conf={getattr(fv,'confidence',None)}")
        items = flds.get("Items")
        n = len(getattr(items, "value_array", []) or []) if items else 0
        print("  line items:", n)
except Exception as exc:
    print("FAILED:", type(exc).__name__, str(exc)[:300])
    sys.exit(1)
