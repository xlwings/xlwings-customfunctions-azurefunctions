import logging
import os
from dataclasses import dataclass
from typing import List

import jwt

# See: https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration
jwks_uri = "https://login.microsoftonline.com/common/discovery/v2.0/keys"


required_roles = os.environ["XLWINGS_REQUIRED_ROLES"].split(",")
azuread_tenant_id = os.environ["AZUREAD_TENANT_ID"]
azuread_client_id = os.environ["AZUREAD_CLIENT_ID"]


@dataclass
class User:
    name: str
    email: str
    roles: List


def authenticate(access_token: str):
    """Returns a tuple: user, error_message"""
    if access_token.lower().startswith("bearer"):
        parts = access_token.split()
        if len(parts) != 2:
            return False, "Invalid token"
        access_token = parts[1]
    else:
        return False, "Token has to start with Bearer"
    jwks_client = jwt.PyJWKClient(jwks_uri)
    key = jwks_client.get_signing_key_from_jwt(access_token)

    token_version = jwt.decode(access_token, options={"verify_signature": False}).get(
        "ver"
    )

    # https://learn.microsoft.com/en-us/azure/active-directory/develop/access-tokens#token-formats
    # Upgrade to 2.0:
    # https://learn.microsoft.com/en-us/answers/questions/639834/how-to-get-access-token-version-20.html
    if token_version == "1.0":
        issuer = f"https://sts.windows.net/{azuread_tenant_id}/"
        audience = f"api://{azuread_client_id}"
    elif token_version == "2.0":
        issuer = f"https://login.microsoftonline.com/{azuread_tenant_id}/v2.0"
        audience = azuread_client_id
    else:
        return (
            None,
            f"Unsupported token version: {token_version}",
        )
    try:
        claims = jwt.decode(
            access_token,
            key.key,
            algorithms=["RS256"],
            issuer=issuer,
            audience=audience,
        )
    except Exception:
        return None, "Couldn't validate access token"

    user = User(
        name=claims["name"],
        email=claims["preferred_username"],
        roles=claims.get("roles") if claims.get("roles") else [],
    )
    if required_roles and not set(required_roles).issubset(user.roles):
        return (
            None,
            f'Required role(s) missing: {", ".join(set(required_roles) - set(user.roles))}',
        )
    else:
        return (
            user,
            None,
        )
