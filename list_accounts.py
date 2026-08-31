"""
list_accounts.py — prints your FreeAgent bank accounts and their IDs, so
you can identify the polluted account and the new clean account for .env.

Usage:
    python list_accounts.py
"""
import freeagent_client as fa


def main():
    accounts = fa.get_bank_accounts()
    print(f"{'ID':<6} {'Name':<35} {'Type':<22} {'Status':<10} {'Balance':>12}")
    print("-" * 92)
    for acc in accounts:
        acc_id = acc["url"].rstrip("/").split("/")[-1]
        print(
            f"{acc_id:<6} {acc.get('name', ''):<35} {acc.get('type', ''):<22} "
            f"{acc.get('status', ''):<10} {acc.get('current_balance', ''):>12}"
        )
    print(
        "\nCopy the IDs for your polluted account and your new clean "
        "account into .env as ACCOUNT_POLLUTED_ID / ACCOUNT_CLEAN_ID."
    )


if __name__ == "__main__":
    main()
