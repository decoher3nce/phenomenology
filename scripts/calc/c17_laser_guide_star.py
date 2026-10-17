"""C17 first calculation: a ground sodium laser guide star seen from LEO.

Question (from data.json, concept C17 "first"): beacon irradiance at 500 km for
a 20 W laser and the SNR for a 30 cm imager at 1 ms exposure.

Method: resonant absorption in the mesospheric sodium layer gives the
fluorescence rate; the beacon radiates roughly isotropically, so the flux at
the satellite follows from the slant range. Background is the nightside ground
and the sodium nightglow through a 1 nm filter. The ground-telescope return is
printed as a sanity check against published LGS returns (order 1e7 ph/s/m^2
for 20 W class lasers). Standard library only.
"""
import math

# ---- assumptions (edit here) ------------------------------------------------
P_laser   = 20.0                # W, CW at 589.16 nm (ESO/Toptica class)
lam       = 589.16e-9           # m
N_col     = 4e13                # m^-2, sodium column abundance (range 2-8e13 over the year)
sigma_D   = 1.0e-15             # m^2, Doppler-broadened D2 cross section near line centre, ~190 K
f_sat     = 0.5                 # saturation / optical-pumping efficiency factor (repumped CW beam, ~1 m spot)
h_layer   = 90e3                # m, sodium layer altitude
h_sat     = 500e3               # m, satellite altitude
T_exp     = 1e-3                # s, exposure (pushbroom line time)
D_ap      = 0.30                # m, imager aperture
eta       = 0.50                # optics throughput x QE at 589 nm
tau_up    = 0.85                # ground-to-layer transmission at 589 nm
filt_bw   = 1.0                 # nm, narrowband filter
L_ground  = 1e-9 * 1e4          # W m^-2 sr^-1, dark nightside land, 400-900 nm (1 nW/cm2/sr); x100 for a city
band_eff  = 500.0               # nm, width over which L_ground is spread
nightglow = 100.0               # Rayleigh, sodium D nightglow (50-200 R)
pix_ifov  = 2e-6                # rad, pixel IFOV (1 m GSD class imager from 500 km)
snr_min   = 5.0
sigma_cal = 0.01                # target radiometric uncertainty (1 %)

h, c = 6.626e-34, 2.998e8
E_ph = h * c / lam
rate_laser = P_laser / E_ph                      # photons / s leaving the laser
tau_opt = N_col * sigma_D                        # optical depth of the sodium layer
rate_fluor = rate_laser * tau_up * (1 - math.exp(-tau_opt)) * f_sat   # photons / s re-emitted
# isotropic emission (circular polarisation gives ~1.5x more along the beam axis; ignored, conservative)
per_sr = rate_fluor / (4 * math.pi)

R_ground = h_layer
R_sat = h_sat - h_layer
flux_ground = per_sr / R_ground ** 2              # photons / s / m^2 at a ground telescope
flux_sat = per_sr / R_sat ** 2                    # photons / s / m^2 at the satellite
A_ap = math.pi * (D_ap / 2) ** 2
sig = flux_sat * A_ap * eta * T_exp               # detected photons per exposure

# background per pixel through the filter
omega_pix = pix_ifov ** 2
A_pix_ground = (pix_ifov * h_sat) ** 2
omega_ap_from_ground = A_ap / h_sat ** 2
bg_ground_rate = (L_ground / band_eff * filt_bw) * A_pix_ground * omega_ap_from_ground / E_ph * eta
bg_glow_rate = (nightglow * 1e10 / (4 * math.pi)) * A_ap * omega_pix * eta   # 1 R = 1e10 ph m^-2 s^-1 column / 4pi sr
bg = (bg_ground_rate + bg_glow_rate) * T_exp
snr = sig / math.sqrt(sig + bg)

print(f"laser photons {rate_laser:.2e}/s; sodium optical depth {tau_opt:.3f}; fluorescence {rate_fluor:.2e}/s")
print(f"sanity: return at a ground telescope {flux_ground:.2e} ph/s/m2 (published 20 W returns: ~1e7)")
print(f"\nat the satellite (slant {R_sat/1e3:.0f} km): {flux_sat:.2e} ph/s/m2")
print(f"detected in {T_exp*1e3:.0f} ms with {D_ap*100:.0f} cm aperture: {sig:.1f} photons")
print(f"background per pixel per exposure: ground {bg_ground_rate*T_exp:.4f}, nightglow {bg_glow_rate*T_exp:.4f} photons")
print(f"SNR per {T_exp*1e3:.0f} ms exposure: {snr:.1f}")

t_cal = (1 / sigma_cal) ** 2 / (sig / T_exp)      # integration for 1 % shot-noise radiometry
print(f"\nfor {sigma_cal*100:.0f} % radiometric precision you need ~{1/sigma_cal**2:.0f} photons = {t_cal:.2f} s of staring")
for D in (0.3, 0.5, 1.0):
    s = flux_sat * math.pi * (D / 2) ** 2 * eta * T_exp
    print(f"  aperture {D*100:.0f} cm: {s:.0f} photons per ms, SNR {s/math.sqrt(s+bg):.0f}; 1 % radiometry in {(1/sigma_cal**2)/(s/T_exp):.2f} s")

if snr >= snr_min:
    v = "link CLOSES: clean point source, negligible background; geolocation and MTF at the line rate, radiometry with a short stare"
elif snr >= 3:
    v = ("link is MARGINAL at 1 ms for this aperture (SNR %.1f): usable for centroiding with TDI or a stare, "
         "comfortable from 50 cm up; radiometry needs a stare of order a second" % snr)
else:
    v = "link FAILS at 1 ms for this aperture; needs TDI, a stare, or a larger aperture"
print("\nVERDICT:", v)
