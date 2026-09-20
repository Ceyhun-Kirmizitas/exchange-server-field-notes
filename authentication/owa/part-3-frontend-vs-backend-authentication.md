# Exchange OWA Frontend vs Backend Authentication: Backend 401 After Successful FBA Login

## Frontend vs Backend Authentication

This lab kept the frontend OWA configuration on the normal FBA baseline and changed only the backend OWA authentication settings.

The goal was to separate two different conditions:

```text
Frontend accepted the user credentials
!=
Backend OWA request completed successfully
```

## Starting point

Frontend:

```text
Default Web Site\owa

Exchange FBA           : Enabled
Basic Authentication   : Enabled
Windows Authentication : Disabled
```

Backend:

```text
Exchange Back End\owa

Anonymous Authentication : Enabled
Windows Authentication   : Enabled
Basic Authentication     : Disabled
```

## Test results

| Test | Backend Anonymous | Backend Windows | Result |
|---|---:|---:|---|
| Baseline | Enabled | Enabled | OWA works |
| Test A | Disabled | Enabled | Tested FBA flow still works |
| Test B | Enabled | Disabled | Frontend accepts FBA, backend returns 401 |

With backend Anonymous disabled but Windows Authentication still enabled, the tested FBA flow completed successfully and the backend IIS log showed the authenticated user with HTTP 200.

With backend Windows Authentication disabled, the browser still showed the normal OWA FBA page and accepted valid credentials. The failure happened after that:

```text
POST /owa/auth.owa  -> 302
GET  /owa/          -> 401
GET  /owa/auth/logon.aspx?...reason=0...
```

HttpProxy made the boundary much clearer:

```text
AuthenticationType : FBA
IsAuthenticated    : true
AuthenticatedUser  : LAB6\auser601
HttpStatus         : 401
BackEndStatus      : 401
ErrorCode          : ProtocolError
RoutingHint        : WindowsIdentity
```

That is the key troubleshooting lesson from this test: a valid frontend sign-in does not prove the backend OWA request succeeded.

## Practical lesson

Troubleshoot OWA in this order:

```text
Did the browser reach the frontend?
        ↓
Did the frontend accept the credentials?
        ↓
Did the authenticated request complete on the backend?
```

The backend defaults were restored after testing. These A/B changes were for lab analysis, not production configuration guidance.

## Full article

[Exchange Server SE / 2019 OWA Authentication Deep Dive — Part 3: Frontend vs Backend Authentication](https://ceyhunkirmizitas.net/exchange-owa-frontend-backend-authentication-part-3/)
