# Exchange Server OWA Authentication Request Flow: Frontend, HttpProxy and Backend

## How OWA Authentication Really Works

OWA authentication is easier to troubleshoot when the frontend, Exchange authentication layer, HttpProxy, and backend are treated as separate parts of the same request path.

## Lab baseline

```text
Browser
  |
  v
Default Web Site\owa        (frontend, HTTPS 443)
  |
  v
Exchange OWA authentication
  |
  v
HttpProxy / routing
  |
  v
Exchange Back End\owa      (backend, HTTPS 444)
  |
  v
OWA application / mailbox services
```

The Exchange frontend baseline used in the lab was:

```text
FormsAuthentication           : True
BasicAuthentication           : True
WindowsAuthentication         : False
DigestAuthentication          : False
InternalAuthenticationMethods : {Basic, Fba}
ExternalAuthenticationMethods : {Fba}
```

The important point is that Exchange OWA Forms-Based Authentication is not the same thing as the generic IIS Forms Authentication feature.

On the frontend OWA application, IIS Basic Authentication is enabled while IIS Forms Authentication remains disabled. On the backend OWA application, Anonymous and Windows Authentication are enabled by default.

## Practical troubleshooting lesson

Do not infer the whole OWA authentication flow from a single Exchange property or IIS authentication page.

I normally separate the evidence into three layers:

```text
Frontend IIS      -> C:\inetpub\logs\LogFiles\W3SVC1
Exchange HttpProxy -> ...\Logging\HttpProxy\Owa
Backend IIS       -> C:\inetpub\logs\LogFiles\W3SVC2
```

A few fields that look definitive are not always enough by themselves. For example, an HTTP 401 can be part of a normal Windows Integrated Authentication challenge, and `WindowsAuthentication=False` on the frontend does not mean Windows Authentication is absent from the backend request path.

## Full article

[Exchange Server SE / 2019 OWA Authentication Deep Dive — Part 1: How OWA Authentication Really Works](https://ceyhunkirmizitas.net/exchange-owa-authentication-deep-dive-part-1/)
