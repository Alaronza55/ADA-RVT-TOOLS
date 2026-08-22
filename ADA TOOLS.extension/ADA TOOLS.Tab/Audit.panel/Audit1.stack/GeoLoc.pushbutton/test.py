import math

def lambert72_to_wgs84(x, y):
    """
    Convert Lambert72 (EPSG:31370) to WGS84 (EPSG:4326)
    Pure Python implementation
    """
    # Lambert72 parameters
    a = 6378388.0  # Semi-major axis
    e = 0.08199188998  # Eccentricity
    
    # Lambert72 projection parameters
    lat0 = math.radians(90.0)
    lon0 = math.radians(4.356939722222222)
    lat1 = math.radians(49.833333333333336)
    lat2 = math.radians(51.166666666666664)
    x0 = 150000.01256
    y0 = 5400088.4378
    
    # Inverse projection
    n = math.log(math.cos(lat1) / math.cos(lat2)) / math.log(
        math.tan(math.pi / 4 + lat2 / 2) / math.tan(math.pi / 4 + lat1 / 2)
    )
    
    F = math.cos(lat1) * math.pow(math.tan(math.pi / 4 + lat1 / 2), n) / n
    
    rho0 = a * F
    
    dx = x - x0
    dy = y - y0
    
    rho = math.sqrt(dx * dx + (rho0 - dy) * (rho0 - dy))
    theta = math.atan2(dx, rho0 - dy)
    
    longitude = theta / n + lon0
    latitude = 2 * math.atan(math.pow(a * F / rho, 1 / n)) - math.pi / 2
    
    # Convert to degrees
    latitude = math.degrees(latitude)
    longitude = math.degrees(longitude)
    
    return latitude, longitude

# Your coordinates
x_lambert = 148320.62
y_lambert = 153448.56

latitude, longitude = lambert72_to_wgs84(x_lambert, y_lambert)

print(f"Original Lambert72 coordinates:")
print(f"X (Easting): {x_lambert}")
print(f"Y (Northing): {y_lambert}")
print(f"\nConverted to WGS84:")
print(f"Latitude: {latitude:.15f}")
print(f"Longitude: {longitude:.15f}")
