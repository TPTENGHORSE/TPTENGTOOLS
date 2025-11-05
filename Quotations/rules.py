from typing import List, Tuple


# Legs legend (for reference in codebase):
# 1 = Origin inland (supplier/plant → POL/Airport)
# 2 = Main carriage (ocean/air)
# 3 = Destination inland (POD/Airport → destination plant)


_INCOTERM_RULES = {
    # Incoterm: (legs_we_pay, type_of_flow)
    # Buyer-perspective mapping: we always compute the legs that the BUYER must pay.
    # Legend: 1=Origin inland, 2=Main carriage (ocean/air), 3=Destination inland.
    # Notes:
    # - Keep type_of_flow="Overseas" whenever any sea leg is relevant or to avoid the Inland override in downstream logic.
    # - For terms where buyer pays no transport (e.g., DAP/DDP), leave legs empty [].
    # Common ocean terms (buyer responsibilities):
    "EXW": ([1, 2, 3], "Overseas"),   # Buyer arranges/pays everything from supplier: legs 1,2,3
    "FCA": ([2, 3], "Overseas"),      # Seller delivers to carrier at origin; buyer pays main carriage + destination inland
    "FOB": ([2, 3], "Overseas"),      # Seller delivers on board at POL; buyer pays main carriage + destination inland
    "CFR": ([3], "Overseas"),         # Seller pays main carriage to POD; buyer pays destination inland
    "CIF": ([3], "Overseas"),         # Same as CFR (insurance by seller); buyer pays destination inland
    "CPT": ([3], "Overseas"),         # Carriage Paid To (by seller); buyer pays destination inland
    "CIP": ([3], "Overseas"),         # Carriage & Insurance Paid (by seller); buyer pays destination inland
    "DAP": ([1], "Inland"),          # Delivered At Place (by seller); buyer pays no transport
    "DDP": ([], "Overseas"),          # Delivered Duty Paid (by seller); buyer pays no transport
}


def flow_by_incoterm(incoterm: str) -> Tuple[List[int], str]:
    """
    Returns (legs_included, type_of_flow) for a given incoterm.

    legs_included: list of ints using the legend above.
    type_of_flow: one of {"Overseas", "Inland"}. (Extendable to "Air", "Combined".)

    Raises ValueError if incoterm is unknown.
    """
    if not incoterm:
        raise ValueError("Incoterm vacío")
    key = str(incoterm).strip().upper()
    if key not in _INCOTERM_RULES:
        raise ValueError(f"Incoterm no soportado: {incoterm}")
    return _INCOTERM_RULES[key]


def is_leg_included(incoterm: str, leg: int) -> bool:
    legs, _ = flow_by_incoterm(incoterm)
    return int(leg) in legs
