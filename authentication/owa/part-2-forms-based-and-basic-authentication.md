# Exchange OWA Forms-Based vs Basic Authentication: HTTP 401, Browser Prompt and IIS Changes

## Forms-Based and Basic Authentication

This lab changed only the frontend OWA authentication settings and compared the Exchange configuration, IIS state, and browser behavior.

## Test matrix

| Test | Frontend configuration | Browser result |
|---|---|---|
| Baseline | FBA + Basic | Exchange OWA sign-in page |
| Test 1 | FBA disabled, no other method enabled | Terminal HTTP 401 |
| Test 2 | Basic only | Browser Basic credential prompt |
| Test 3 | FBA restored | Exchange OWA sign-in page |

## Key findings

Disabling Forms-Based Authentication without enabling another frontend method left the browser with no usable authentication path and produced a final HTTP 401.

Enabling Basic Authentication while FBA remained disabled changed the user experience completely:

```text
FBA enabled
-> Exchange OWA sign-in page

FBA disabled + Basic enabled
-> Browser HTTP Basic credential prompt
```

Re-enabling FBA restored the baseline:

```text
FormsAuthentication           : True
BasicAuthentication           : True
WindowsAuthentication         : False
DigestAuthentication          : False
InternalAuthenticationMethods : {Basic, Fba}
```

Another useful lab observation was that `Set-OwaVirtualDirectory` changed the relevant frontend IIS authentication state as well as the Exchange properties.

## Troubleshooting lesson

Separate these questions:

```text
What is configured?
-> Get-OwaVirtualDirectory

What is IIS allowing?
-> IIS Authentication settings

What did the request actually do?
-> IIS + HttpProxy logs
```

`ExternalAuthenticationMethods={Fba}` did not reliably describe the runtime method used by a specific request during these tests.

## Full article

[Exchange Server SE / 2019 OWA Authentication Deep Dive — Part 2: Forms-Based and Basic Authentication](https://ceyhunkirmizitas.net/exchange-owa-fba-basic-authentication-part-2/)
