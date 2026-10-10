"""Exercise certificate and hostname verification through the native trust store."""

from contextlib import contextmanager
import shutil
import socket
import ssl
import subprocess
import threading

import pytest

from pastemd.utils.https import build_https_opener, create_https_context


@pytest.fixture
def tls_identity(tmp_path):
    openssl = shutil.which("openssl")
    if not openssl:
        pytest.skip("openssl is required to generate an untrusted local test certificate")
    cert, key = tmp_path / "test-cert.pem", tmp_path / "test-key.pem"
    ca, ca_key, csr = tmp_path / "ca.pem", tmp_path / "ca-key.pem", tmp_path / "server.csr"
    extensions = tmp_path / "server.ext"
    extensions.write_text("basicConstraints=critical,CA:FALSE\n"
                          "keyUsage=critical,digitalSignature,keyEncipherment\n"
                          "extendedKeyUsage=serverAuth\nsubjectAltName=DNS:localhost\n")

    def run(*args):
        subprocess.run([openssl, *args], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)

    run(
        "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "1",
        "-subj", "/CN=PasteMD temporary test CA", "-addext", "basicConstraints=critical,CA:TRUE",
        "-addext", "keyUsage=critical,keyCertSign,cRLSign", "-out", str(ca), "-keyout", str(ca_key),
    )
    run(
        "req", "-new", "-newkey", "rsa:2048", "-nodes", "-subj", "/CN=localhost",
        "-out", str(csr), "-keyout", str(key),
    )
    run(
        "x509", "-req", "-in", str(csr), "-CA", str(ca), "-CAkey", str(ca_key),
        "-CAcreateserial", "-days", "1", "-sha256", "-extfile", str(extensions), "-out", str(cert),
    )
    return cert, key, ca


@contextmanager
def local_tls_server(identity):
    cert, key, _ = identity
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(cert, key)
    listener = socket.socket()
    listener.bind(("127.0.0.1", 0))
    listener.listen(1)
    listener.settimeout(5)
    errors = []

    def serve():
        try:
            conn, _ = listener.accept()
            with conn:
                conn.settimeout(5)
                try:
                    with context.wrap_socket(conn, server_side=True) as secured:
                        secured.recv(1)
                except ssl.SSLError:
                    pass  # A client rejecting the test certificate is expected.
        except Exception as exc:
            errors.append(exc)

    thread = threading.Thread(target=serve, daemon=True)
    thread.start()
    try:
        yield listener.getsockname()
    finally:
        thread.join(timeout=6)
        listener.close()
        assert not thread.is_alive()
        assert not errors


@pytest.mark.parametrize("trusted,hostname,accepted", [
    (False, "localhost", False),
    (True, "localhost", True),
    (True, "wrong-host.invalid", False),
])
def test_native_tls_rejects_untrusted_chain_and_wrong_hostname(tls_identity, trusted, hostname, accepted):
    context = create_https_context()
    assert context.verify_mode == ssl.CERT_REQUIRED
    assert context.check_hostname
    if trusted:
        # Trust this temporary CA in this context, without changing OS trust.
        context.load_verify_locations(cafile=str(tls_identity[2]))
    with local_tls_server(tls_identity) as address:
        with socket.create_connection(address, timeout=5) as conn:
            if accepted:
                with context.wrap_socket(conn, server_hostname=hostname) as secured:
                    secured.send(b"x")
            else:
                with pytest.raises(ssl.SSLCertVerificationError):
                    context.wrap_socket(conn, server_hostname=hostname)


@pytest.mark.parametrize("use_proxy", [False, True])
def test_both_transports_keep_verified_native_tls(use_proxy):
    import urllib.request

    opener = build_https_opener(use_proxy)
    handler = next(h for h in opener.handlers if isinstance(h, urllib.request.HTTPSHandler))
    assert handler._context.verify_mode == ssl.CERT_REQUIRED
    assert handler._context.check_hostname
    if not use_proxy:
        assert not any(isinstance(h, urllib.request.ProxyHandler) for h in opener.handlers)
