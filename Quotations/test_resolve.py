from Quotations.generate_quote import QTOOL_DIR
try:
    from Quotations.Distances import GeoIndex, resolve_point
except Exception:
    from Quotations.distances import GeoIndex, resolve_point

geo = GeoIndex.load_from_dir(QTOOL_DIR)
cc = 'IN'
print('GeoIndex loaded from', QTOOL_DIR)
for city, zipc in [('Mundhwa','34190'), ('Mundhwa',''), ('Pune',''), ('Pune','411036')]:
    lat, lon, src = resolve_point(geo, cc, city=city, zip_code=zipc)
    print(f'city={city!r} zip={zipc!r} ->', lat, lon, src)
