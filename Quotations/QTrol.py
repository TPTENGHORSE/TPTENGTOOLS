"""
QTrol - core entry point (skeleton)

This module currently exposes CLI helpers to resolve flow settings from an
Incoterm, based on the rules provided:

Incoterm  | Legs             | Type of flow
----------|------------------|-------------
FCA       | leg 1 + 2 + 3    | Overseas
DAP       | leg 1 + 2 + 3    | Inland
FOB       | leg 2 + 3        | Overseas
CIF       | leg 3            | Overseas

Legs legend:
 1 = Origin inland (supplier/plant → POL/Airport)
 2 = Main carriage (ocean/air)
 3 = Destination inland (POD/Airport → destination plant)

Future: Extend type_of_flow to Air/Combined when additional criteria is provided.
"""

from typing import List
import argparse
from .rules import flow_by_incoterm


def describe_flow(incoterm: str) -> str:
	legs, flow_type = flow_by_incoterm(incoterm)
	legs_str = "+".join(str(x) for x in legs)
	return f"Incoterm {incoterm.upper()}: legs {legs_str} | type_of_flow={flow_type}"


def main(argv: List[str] | None = None) -> int:
	parser = argparse.ArgumentParser(description="QTrol: flow resolver by Incoterm")
	parser.add_argument("incoterm", help="Incoterm code (FCA, DAP, FOB, CIF)")
	args = parser.parse_args(argv)
	try:
		print(describe_flow(args.incoterm))
		return 0
	except Exception as e:
		print(f"ERROR: {e}")
		return 2


if __name__ == "__main__":
	raise SystemExit(main())
