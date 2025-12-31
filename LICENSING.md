# Licensing & Billing (Advanced Pie Chart Pro)

This visual uses Power BI’s built-in **Visual Licensing** API (`host.licenseManager`) to enable/disable Pro-only features.

## What you must set up in Microsoft Partner Center

1. **Create a separate offer** for the Pro visual (recommended)
   - Treat Pro as a separate AppSource/Marketplace submission with its own `guid` and `name`.

2. **Enable Microsoft Licensing & Billing** for the offer
   - Configure the offer as **paid** (or with paid plans) using Partner Center’s licensing/billing options.

3. **Create a Service Plan identifier (important)**
   - In Partner Center licensing configuration, create at least one **service plan**.
   - Set the Service Plan Identifier to match what the visual checks in code.

   Current code expects this plan identifier:
   - `advancedPieChartPro`

   If you use a different identifier in Partner Center, update the code constant:
   - `proServicePlanIds` in `src/visual.ts`

4. Publish/update the offer
   - After publishing, Power BI will provide license plan information to the visual at runtime.

## How the visual enforces licensing

- The visual calls:
  - `host.licenseManager.getAvailableServicePlans()`
- If it finds a matching plan with state:
  - `Active` or `Warning`
  then Pro features are enabled.
- If the user is not licensed, the visual:
  - blocks the Pro feature
  - calls `notifyLicenseRequired(...)` and `notifyFeatureBlocked(...)` so the user sees an upgrade prompt.

## Pro features currently gated

- **Top N** (shows exactly the top N slices) (Format pane → **Pro features**)
- **Min % → Others** grouping (Format pane → **Pro features**)

## Dev/Test notes

- For development right now, licensing enforcement is **OFF** so Pro features can be tested without an active plan.
   - Toggle this in `src/visual.ts`: set `enforceLicensing` to `true` when you’re ready to go live.
- In environments where licensing is unsupported, the visual treats Pro as not licensed.
- You can still package and run the visual; Pro features will just remain disabled until the offer/licensing is configured.
