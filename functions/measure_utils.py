# functions/measure_utils.py
import math

def calculate_height(ruler_rect, piquet_rect, ruler_height_cm, fov_deg=0, image_width_px=0):
    try:
        ruler_px = ruler_rect.height()
        piquet_px = piquet_rect.height()
        if ruler_px <= 0:
            raise ValueError("La hauteur en pixels de la règle doit être positive.")
        if fov_deg == 0 or image_width_px == 0:
            ratio = ruler_height_cm / ruler_px
            return piquet_px * ratio
        fov_rad = math.radians(fov_deg)
        angle_per_pixel = fov_rad / image_width_px
        ruler_angle = ruler_px * angle_per_pixel
        if ruler_angle == 0:
            return None
        distance = (ruler_height_cm / 100.0 / 2) / math.tan(ruler_angle / 2)
        piquet_angle = piquet_px * angle_per_pixel
        piquet_height_m = 2 * distance * math.tan(piquet_angle / 2)
        return piquet_height_m * 100.0
    except Exception as e:
        print(f"Erreur lors du calcul de la hauteur : {e}")
        return None
