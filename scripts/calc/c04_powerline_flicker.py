"""C04 first calculation: photon budget for mains-frequency flicker seen from LEO.

Question (from data.json, concept C04 "first"): photon rate for a 100 m ground
pixel at typical urban radiance with a 30 cm aperture, and the SNR at 120 Hz
over a 1 s dwell.

Method: radiometry for the mean photon rate, then the matched-filter SNR of a
sinusoidal modulation on a Poisson photon stream, cross-checked by a Monte Carlo
of the same case. Standard library only. Run: python3 c04_powerline_flicker.py
"""
import math, random

# ---- assumptions (edit here) ------------------------------------------------
L_urban   = 100e-9 * 1e4        # W m^-2 sr^-1: 100 nW cm^-2 sr^-1, bright urban core (VIIRS DNB class)
L_suburb  = 10e-9 * 1e4         # W m^-2 sr^-1: 10 nW cm^-2 sr^-1, suburban / small town
lam       = 600e-9              # m, effective wavelength of the lamp mix (DNB-like 500-900 nm band)
gsd       = 100.0               # m, ground pixel side
D_ap      = 0.30                # m, aperture diameter
R         = 500e3               # m, nadir range from LEO
tau_atm   = 0.80                # one-way atmospheric transmission, clear night
eta       = 0.50                # optics throughput x detector QE
m_eff     = 0.10                # effective modulation depth at 2x mains after mixing lamp types (LED 5-30 %, HID 10-40 %, incandescent ~1 %)
f_mod     = 120.0               # Hz, 2 x 60 Hz
T_dwell   = 1.0                 # s, stare time on the pixel (needs body pointing or back-scan; ground track moves 7 km/s)
f_samp    = 1000.0              # Hz, frame rate
snr_min   = 5.0                 # detection threshold on the 120 Hz Fourier amplitude

h, c = 6.626e-34, 2.998e8
E_ph = h * c / lam
A_ground = gsd ** 2
omega_ap = math.pi * (D_ap / 2) ** 2 / R ** 2       # sr subtended by the aperture from the ground

def budget(L):
    P = L * A_ground * omega_ap * tau_atm            # W collected
    N = P / E_ph * eta                               # detected photons / s
    # Matched-filter SNR for relative modulation m on a Poisson stream over T:
    # signal amplitude in the f_mod bin = m N T / 2, shot-noise std per bin = sqrt(N T / 2)
    snr = m_eff * math.sqrt(N * T_dwell / 2)
    return P, N, snr

def monte_carlo(N, seed=1):
    random.seed(seed)
    n = int(T_dwell * f_samp); dt = 1 / f_samp
    k = int(round(f_mod * T_dwell))                  # Fourier bin of 120 Hz
    xs = []
    for i in range(n):
        mu = N * dt * (1 + m_eff * math.sin(2 * math.pi * f_mod * i * dt))
        # Poisson via normal approximation (mu >> 1 here)
        xs.append(random.gauss(mu, math.sqrt(mu)))
    def amp(bin_):
        re = sum(x * math.cos(2 * math.pi * bin_ * i / n) for i, x in enumerate(xs))
        im = sum(x * math.sin(2 * math.pi * bin_ * i / n) for i, x in enumerate(xs))
        return math.hypot(re, im)
    sig = amp(k)
    noise = sorted(amp(b) for b in range(k + 20, k + 60))   # empty bins above the signal
    noise_rms = math.sqrt(sum(v * v for v in noise) / len(noise))
    return sig / noise_rms

print(f"aperture solid angle from ground: {omega_ap:.2e} sr; photon energy {E_ph:.2e} J")
for label, L in (("bright urban", L_urban), ("suburban", L_suburb)):
    P, N, snr = budget(L)
    print(f"\n{label}: radiance {L*1e5:.0f} nW/cm2/sr")
    print(f"  collected power     {P:.2e} W")
    print(f"  detected photons    {N:.2e} /s   = {N/f_samp:.0f} per {1e3/f_samp:.0f} ms frame")
    print(f"  analytic SNR at {f_mod:.0f} Hz, m={m_eff}, T={T_dwell}s:  {snr:.0f}")
    print(f"  Monte Carlo SNR:    {monte_carlo(N):.0f}")
    m_min = snr_min / math.sqrt(N * T_dwell / 2)
    print(f"  smallest detectable modulation depth at SNR {snr_min:.0f}: {m_min*100:.2f} %")

verdict = budget(L_suburb)[2] >= snr_min
print("\nVERDICT:", "photon budget CLOSES with wide margin" if verdict else "photon budget FAILS",
      "- the limit is not photons but the platform: a 1 s stare on a 100 m pixel from LEO needs",
      "body pointing or back-scan, and pointing jitter must be quiet near 100-120 Hz.")
