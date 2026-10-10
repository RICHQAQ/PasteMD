"""HTTPS transports using the operating system's certificate trust store."""

import urllib.request

from .logging import log


def create_https_context():
    # Keep this import after VersionChecker's Windows OpenSSL diagnostics/setup.
    import ssl
    import truststore

    # macOS: Security framework; Windows: CryptoAPI; Linux: OpenSSL.
    # Do not depend on the Python installation used by the build machine, or
    # inject global SSL replacements into unrelated libraries.
    context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    log("[update] TLS trust: operating system certificate store; "
        f"certificate verification={context.verify_mode.name}, "
        f"hostname verification={context.check_hostname}")
    return context


def build_https_opener(use_proxy: bool):
    handlers = [urllib.request.HTTPSHandler(context=create_https_context())]
    if not use_proxy:
        handlers.insert(0, urllib.request.ProxyHandler({}))
    return urllib.request.build_opener(*handlers)
