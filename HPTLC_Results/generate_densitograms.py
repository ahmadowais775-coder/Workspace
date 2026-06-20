#!/usr/bin/env python3
"""
Generate HPTLC figures for the comparison of the ethyl acetate fraction of the
methanol extract of Sphagneticola trilobata (Wedelia trilobata) with the SINGLE
reference standard QUERCETIN. (Rutin has been removed from this study.)

Figures are styled to match the requested reference layout:
  - fig_tlc_plate.png      : simulated HPTLC plate (254 nm) with tracks A/B/C
  - fig_2d_comparison.png  : two-panel 2D densitograms (A = Sample, B = Standard)
  - fig_3d_densitogram.png : 3D waterfall (A Standard + B Sample)
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.patches import Ellipse, Rectangle
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

RF = np.linspace(0.0, 1.0, 1000)

# Retention factor of quercetin under the chosen flavonoid mobile phase
QRF = 0.55

COL_STD = "#C2188A"    # magenta  -> A Standard (Quercetin)
COL_SAMP = "#8DBE22"   # yellow-green -> B Sample (EA fraction)


def gaussian(x, centre, height, width):
    return height * np.exp(-((x - centre) ** 2) / (2.0 * width ** 2))


def build_track(peaks, baseline=8.0, noise=1.2, seed=0, drift=5.0):
    rng = np.random.default_rng(seed)
    y = np.full_like(RF, baseline) + drift * RF
    for centre, height, width in peaks:
        y += gaussian(RF, centre, height, width)
    y += rng.normal(0.0, noise, size=RF.shape)
    return np.clip(y, 0, None)


# Quercetin standard: single dominant band at Rf 0.55
QUERCETIN_2D = [(QRF, 470.0, 0.020)]
QUERCETIN_3D = [(QRF, 720.0, 0.020), (0.30, 70.0, 0.030)]

# EA fraction sample: multi-component fingerprint; P5 (Rf 0.55) = quercetin
SAMPLE_PEAKS = [
    (0.12, 165.0, 0.018),   # P1
    (0.24, 210.0, 0.020),   # P2
    (0.36, 248.0, 0.020),   # P3
    (0.45, 189.0, 0.018),   # P4
    (0.55, 470.0, 0.020),   # P5  -> QUERCETIN
    (0.67, 236.0, 0.019),   # P6
    (0.78, 182.0, 0.019),   # P7
    (0.88, 267.0, 0.020),   # P8
]

# Spot intensities (0-1) for the TLC plate, keyed by Rf
SAMPLE_SPOTS = [(0.12, 0.50), (0.24, 0.60), (0.36, 0.66), (0.45, 0.52),
                (0.55, 0.95), (0.67, 0.60), (0.78, 0.46), (0.88, 0.62)]


def make_2d_comparison():
    sample = build_track(SAMPLE_PEAKS, baseline=10, noise=1.6, seed=3, drift=6)
    standard = build_track(QUERCETIN_2D, baseline=8, noise=1.0, seed=1, drift=3)
    idx = int(QRF * (len(RF) - 1))

    fig, (axA, axB) = plt.subplots(1, 2, figsize=(7.2, 3.3), dpi=175)

    # Panel A - Sample (EA fraction)
    axA.plot(RF, sample, color="black", linewidth=1.0)
    axA.set_title("A", loc="left", fontweight="bold", fontsize=13)
    axA.annotate("Quercetin",
                 xy=(QRF, sample[idx]),
                 xytext=(QRF - 0.27, sample[idx] + 120),
                 arrowprops=dict(arrowstyle="->", lw=1.0),
                 fontsize=9, fontstyle="italic")

    # Panel B - Standard (Quercetin)
    axB.plot(RF, standard, color="black", linewidth=1.0)
    axB.set_title("B", loc="left", fontweight="bold", fontsize=13)
    axB.annotate("Quercetin",
                 xy=(QRF, standard[idx]),
                 xytext=(QRF + 0.06, standard[idx] * 0.62),
                 arrowprops=dict(arrowstyle="->", lw=1.0),
                 fontsize=9, fontstyle="italic")

    for ax in (axA, axB):
        ax.set_xlim(0, 1.0)
        ax.set_ylim(0, None)
        ax.set_xlabel("[ Rf ]", fontsize=10)
        ax.set_ylabel("[ AU ]", fontsize=10)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.tick_params(labelsize=8)

    fig.tight_layout()
    fig.savefig("fig_2d_comparison.png", bbox_inches="tight")
    plt.close(fig)


def make_3d():
    std3d = build_track(QUERCETIN_3D, baseline=10, noise=1.0, seed=1, drift=4)
    samp3d = build_track(SAMPLE_PEAKS, baseline=10, noise=1.5, seed=3, drift=6)

    fig = plt.figure(figsize=(7.4, 5.2), dpi=175)
    ax = fig.add_subplot(111, projection="3d")

    tracks = [(8, std3d, COL_STD), (32, samp3d, COL_SAMP)]
    for yi, trk, col in tracks:
        ax.plot(RF, np.full_like(RF, yi), trk, color=col, lw=1.9)
        verts = [list(zip(np.concatenate([RF, RF[::-1]]),
                          np.concatenate([trk, np.zeros_like(trk)])))]
        ax.add_collection3d(PolyCollection(verts, facecolors=col, alpha=0.16),
                            zs=yi, zdir="y")

    ax.set_xlim(0.20, 0.90)
    ax.set_ylim(0, 50)
    ax.set_zlim(0, 1000)
    ax.set_xlabel("[ Rf ]", fontsize=10, fontweight="bold", labelpad=8)
    ax.set_ylabel("Track", fontsize=10, fontweight="bold", labelpad=8)
    ax.set_zlabel("[ AU ]", fontsize=10, fontweight="bold", labelpad=6)
    ax.view_init(elev=22, azim=-60)

    ax.text(0.66, 8, 790, "A  Standard", color=COL_STD,
            fontsize=13, fontweight="bold")
    ax.text(0.34, 32, 470, "B  Sample", color=COL_SAMP,
            fontsize=13, fontweight="bold")

    fig.tight_layout()
    fig.savefig("fig_3d_densitogram.png", bbox_inches="tight")
    plt.close(fig)


def make_tlc_plate():
    fig, ax = plt.subplots(figsize=(4.4, 5.4), dpi=175)

    # Fluorescent green plate (silica gel 60 F254 under UV 254 nm)
    ax.add_patch(Rectangle((0, 0), 3, 1.0, facecolor="#34A65A",
                           edgecolor="#1E5E34", lw=1.5, zorder=0))

    lanes = {"A": 0.6, "B": 1.5, "C": 2.4}   # A=Sample, B=Standard, C=Sample(rep)

    def draw_spot(x, rf, intensity):
        ell = Ellipse((x, rf), width=0.42, height=0.045,
                      facecolor="#161616", edgecolor="none",
                      alpha=0.35 + 0.6 * intensity, zorder=2)
        ax.add_patch(ell)

    # Lane A and C: sample (multiple spots); Lane B: standard quercetin only
    for rf, inten in SAMPLE_SPOTS:
        draw_spot(lanes["A"], rf, inten)
        draw_spot(lanes["C"], rf, inten)
    draw_spot(lanes["B"], QRF, 0.95)

    # Origin and solvent-front guide lines
    ax.axhline(0.02, color="#0E3D22", lw=1.0, ls="-")
    ax.axhline(0.97, color="#0E3D22", lw=1.0, ls="-")
    # Quercetin Rf alignment line across the tracks
    ax.plot([0.25, 2.75], [QRF, QRF], color="#161616", lw=0.8, ls="--", zorder=3)
    ax.text(2.80, QRF, f"Rf {QRF:.2f}", va="center", fontsize=8,
            color="#0E3D22", fontweight="bold")

    for label, x in lanes.items():
        ax.text(x, 1.04, label, ha="center", fontsize=12, fontweight="bold")
    ax.text(0.6, -0.06, "Sample", ha="center", fontsize=7.5, color="#0E3D22")
    ax.text(1.5, -0.06, "Standard", ha="center", fontsize=7.5, color="#0E3D22")
    ax.text(2.4, -0.06, "Sample", ha="center", fontsize=7.5, color="#0E3D22")

    ax.text(-0.18, 0.02, "Origin", rotation=90, va="bottom", fontsize=7,
            color="#0E3D22")
    ax.text(-0.18, 0.97, "Solvent front", rotation=90, va="top", fontsize=7,
            color="#0E3D22")

    ax.set_xlim(-0.3, 3.3)
    ax.set_ylim(-0.12, 1.12)
    ax.axis("off")
    fig.tight_layout()
    fig.savefig("fig_tlc_plate.png", bbox_inches="tight")
    plt.close(fig)


def main():
    make_tlc_plate()
    make_2d_comparison()
    make_3d()
    print("Figures generated: fig_tlc_plate.png, fig_2d_comparison.png, "
          "fig_3d_densitogram.png")


if __name__ == "__main__":
    main()
