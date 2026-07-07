"""Third-party analytics integration for the hosted Streamlit app."""

from __future__ import annotations

import streamlit.components.v1 as components

CLARITY_PROJECT_ID = "x1tq0msm2n"
GA4_MEASUREMENT_ID = "G-B9DJEVW7SW"


def render_analytics() -> None:
    """Load hosted-app analytics without adding visible interface content."""
    components.html(
        f"""
        <script async src="https://www.googletagmanager.com/gtag/js?id={GA4_MEASUREMENT_ID}"></script>
        <script>
            window.dataLayer = window.dataLayer || [];
            function gtag(){{dataLayer.push(arguments);}}

            const ga4PageLocation = window.location.href;
            const ga4PagePath = window.location.pathname + window.location.search + window.location.hash;

            if (!window.__excelPixelArtGa4Loaded) {{
                window.__excelPixelArtGa4Loaded = true;
                gtag("js", new Date());
                gtag("config", "{GA4_MEASUREMENT_ID}", {{
                    page_location: ga4PageLocation,
                    page_path: ga4PagePath
                }});
            }} else if (window.__excelPixelArtLastGa4Page !== ga4PageLocation) {{
                gtag("event", "page_view", {{
                    page_location: ga4PageLocation,
                    page_path: ga4PagePath
                }});
            }}
            window.__excelPixelArtLastGa4Page = ga4PageLocation;
        </script>
        <script type="text/javascript">
            if (!window.__excelPixelArtClarityLoaded) {{
                window.__excelPixelArtClarityLoaded = true;
                (function(c,l,a,r,i,t,y){{
                    c[a]=c[a]||function(){{(c[a].q=c[a].q||[]).push(arguments)}};
                    t=l.createElement(r);t.async=1;t.src="https://www.clarity.ms/tag/"+i;
                    y=l.getElementsByTagName(r)[0];y.parentNode.insertBefore(t,y);
                }})(window, document, "clarity", "script", "{CLARITY_PROJECT_ID}");
            }}
        </script>
        """,
        height=0,
        width=0,
    )


def render_clarity_analytics() -> None:
    """Backward-compatible wrapper for the app's analytics injection."""
    render_analytics()
