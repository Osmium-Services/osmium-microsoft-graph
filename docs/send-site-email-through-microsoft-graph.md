---
title: Send your site's email through Microsoft 365
tags: email, microsoft-graph, mail provider, smtp, sending, deliverability
page: admin/settings/microsoft-graph
---

**Microsoft 365** sends your site's email, such as order confirmations and sign-in codes, through Microsoft 365's service instead of SMTP. It is useful where your hosting blocks outbound SMTP.

## Set it up

1. In your Microsoft 365 account, get the details below.
2. Open **Microsoft 365 Mail Settings** and enter the **Tenant ID**, **Client ID** and **Client Secret** from your Azure AD app registration.
3. Press **Save Settings**.
4. Open the **Email** settings page and choose **Microsoft 365** as the **Provider**. The **From Address** and the **Send test email to** button are on that page.

Mail is sent as the From address. Use this when your Microsoft 365 account has SMTP AUTH switched off.

Use **Run diagnostics** on the Email settings page if mail does not arrive.
