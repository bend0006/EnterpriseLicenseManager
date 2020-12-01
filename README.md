# EnterpriseLicenseManager

**Use with caution. This is manipulating settings in the Google Admin Console.**

A script to automate distribution of Google Enterprise Licenses based on the users description.

By adding triggers to the script, it can be set to auto-run e.g. nightly

If the description contains certain keywords, a license will be added.
If the user is suspendes or not qualified for a license, it will be removed.

The script is an extension for a Google Sheets Document.
It relies on a OU-structure like this:
/
/School Name/
/School Name/Staff

The script will look for a sheet named "Schools" containing the OU School Names in the first row.
It will also look for a similar sheet named "Schools not ready" for the "AddLicenses" function.
