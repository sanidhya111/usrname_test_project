# property_scan.py
"""
Utility to scan common Outlook/Exchange user properties for the
currently logged-in user. Safe to run on Windows with Outlook/Exchange.
"""

import sys
import traceback

# ---- Optional Windows/COM guard ----
try:
    import win32com.client as win32
    import pythoncom
    WINDOWS_COM = True
except Exception:
    win32 = None
    pythoncom = None
    WINDOWS_COM = False


def scan():
    if not WINDOWS_COM:
        print("⚠️  Windows COM / pywin32 not available. This tool must be run on Windows with Outlook installed.")
        sys.exit(1)

    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        user = ns.CurrentUser.AddressEntry.GetExchangeUser()
        if not user:
            print("❌ Could not resolve Exchange user for CurrentUser.")
            return

        print(f"Scanning properties for: {user.Name}")
        print("-" * 60)

        # Common, friendly properties
        def safe_attr(obj, name, label=None):
            label = label or name
            try:
                return f"{label}: {getattr(obj, name) or 'Not Set'}"
            except Exception:
                return f"{label}: <error>"

        print(safe_attr(user, "Alias", "Alias"))
        print(safe_attr(user, "JobTitle", "Job Title"))
        print(safe_attr(user, "OfficeLocation", "Office"))
        print(safe_attr(user, "Department", "Department"))
        print(safe_attr(user, "PrimarySmtpAddress", "Email"))
        print(safe_attr(user, "BusinessTelephoneNumber", "Phone"))

        print("-" * 60)
        print("Attempting to read OrganizationalIDNumber (MAPI 0x3A10, PT_UNICODE) ...")

        # MAPI proptag for PR_ORGANIZATIONAL_ID_NUMBER_W == 0x3A10001F
        ORG_ID_TAG = "http://schemas.microsoft.com/mapi/proptag/0x3A10001F"
        try:
            value = user.PropertyAccessor.GetProperty(ORG_ID_TAG)
            print(f"OrganizationalIDNumber: {value}")
        except Exception:
            print("OrganizationalIDNumber: Not found")

        print("-" * 60)
        print("Done.")

    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    scan()
