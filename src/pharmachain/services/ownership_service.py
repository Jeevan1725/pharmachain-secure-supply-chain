"""
PharmaChain - Ownership Transfer Service (Phase 12 refactored).

Before (v1): monolithic, SQL injection, no auth, no signature, no audit.
After  (v2): layered, RBAC, signature verification, hash-chained ledger,
             immutable audit log, testable.
"""
from dataclasses import dataclass


@dataclass
class TransferResult:
    success: bool
    tx_hash: str = ""
    reason: str = ""


class ProductNotFound(Exception):
    pass


class UnauthorizedError(Exception):
    pass


class SignatureError(Exception):
    pass


class OwnershipService:
    def __init__(self, repo, ledger, authz, signer, audit):
        self.repo = repo
        self.ledger = ledger
        self.authz = authz
        self.signer = signer
        self.audit = audit

    def transfer(self, product_id, new_owner, actor, signature):
        if not self.authz.can(actor, "TRANSFER_OWNERSHIP", product_id):
            self.audit.log("TRANSFER_DENIED",
                           {"actor": actor.id, "product": product_id})
            raise UnauthorizedError("Not permitted to transfer this product")

        if not self.signer.verify(actor, product_id, new_owner, signature):
            self.audit.log("TRANSFER_BAD_SIG",
                           {"actor": actor.id, "product": product_id})
            raise SignatureError("Invalid signature")

        product = self.repo.get(product_id)
        if product is None:
            raise ProductNotFound(product_id)

        with self.repo.transaction():
            product.owner = new_owner
            self.repo.save(product)
            tx_hash = self.ledger.append({
                "type": "OWNERSHIP_TRANSFER",
                "product_id": product_id,
                "from": actor.org_id,
                "to": new_owner,
                "signature": signature.hex(),
            })

        self.audit.log("TRANSFER", {
            "actor": actor.id,
            "product_id": product_id,
            "to": new_owner,
            "tx_hash": tx_hash,
        })

        return TransferResult(True, tx_hash=tx_hash)