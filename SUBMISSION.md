# Marketplace/AppSource submission checklist (Power BI custom visual)

This file is a practical checklist to submit the packaged `.pbiviz`.

## 1) Create/prepare public support URLs

Marketplace review generally expects a working public support URL.

- Create GitHub repo: `https://github.com/dpandeycse/advancedPieChart`
- Enable Issues
- Ensure these URLs open publicly:
  - Support URL: https://github.com/dpandeycse/advancedPieChart/issues
  - Repo URL: https://github.com/dpandeycse/advancedPieChart

## 2) Host privacy statement

You can host `PRIVACY.md` on GitHub (repo) and use the rendered page link.

- Privacy statement source: `PRIVACY.md`
- Support contact: dpandeycse@gmail.com

## 3) Validate package

From the `advancedPieChart` folder:

- Install: `npm ci`
- Lint: `npm run lint`
- Package: `npm run package`

Packaged file output:
- `dist/advancedPieChartE8563A8269314CC3ABC3E6FDA9DAE531.1.0.0.0.pbiviz`

## 4) Prepare listing content

You will need:
- Visual name: Advanced Pie Chart
- Short/long description (what it does + key features)
- Support email: dpandeycse@gmail.com
- Support URL (public)
- Privacy policy URL (public)
- Icon/screenshots

## 5) Submit

Submit through Microsoft’s Power BI visuals submission/AppSource flow (Partner Center).

Note: Some features are “recommended” (keyboard navigation, localization, etc.). They may not be strictly required for listing, but can be requested during review depending on the program level.
