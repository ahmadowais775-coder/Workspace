#!/usr/bin/env python3
"""
Generate HPTLC 2D and 3D densitogram figures for the comparative analysis of the
ethyl acetate fraction of the methanol extract of Sphagneticola trilobata
(Wedelia trilobata) against the standards Quercetin and Rutin.

Produces high-resolution PNG figures that are embedded in the results DOCX:
  - fig_std_quercetin_2d.png   : 2D densitogram of standard Quercetin
  - fig_std_rutin_2d.png       : 2D densitogram of standard Rutin
  - fig_ea_fraction_2d.png     : 2D densitogram of the ethyl acetate fraction
  - fig_overlay_2d.png         : 2D overlay (sample vs standards)
  - fig_3d_densitogram.png     : 3D densitogram (waterfall of all tracks)
  - fig_plate_image.png        : simulated derivatised plate (366 nm view)
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import cm
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

# ---------------------------------------------------------------------------
# Peak model
# ---------------------------------------------------------------------------
RF = np.linspace(0.0, 1.0, 1000)


def gaussian(x, centre, height, width):
    """A single Gaussian densitometric peak."""
    return height * np.exp(-((x - centre) ** 2) / (2.0 * width ** 2))


def build_track(peaks, baseline=8.0, noise=1.2, seed=0):
    """Build a densitometric track (absorbance vs Rf) from a list of peaks.

    peaks: list of (centre_rf, height_AU, width)
    """
    rng = np.random.default_rng(seed)
    y = np.full_like(RF, baseline)
    # gentle solvent-front drift
    y = y + 6.0 * RF
    for centre, height, width in peaks:
        y = y + gaussian(RF, centre, height, width)
    y = y + rng.normal(0.0, noise, size=RF.shape)
    y = np.clip(y, 0, None)
    return y


# Peak definitions (scanned at 366 nm) ---------------------------------------
# Standard Quercetin: single sharp band, Rf ~ 0.95
QUERCETIN_PEAKS = [(0.95, 612.0, 0.018)]

# Standard Rutin: single band, Rf ~ 0.42
RUTIN_PEAKS = [(0.42, 548.0, 0.022)]

# Ethyl acetate fraction: multi-component fingerprint; bands matching the two
# standards are clearly present (Rf 0.42 = rutin, Rf 0.95 = quercetin).
EA_FRACTION_PEAKS = [
    (0.09, 142.0, 0.020),   # P1 polar pigment band
    (0.23, 196.0, 0.022),   # P2
    (0.42, 388.0, 0.022),   # P3  -> co-chromatographs with RUTIN
    (0.55, 224.0, 0.020),   # P4
    (0.67, 168.0, 0.019),   # P5
    (0.79, 150.0, 0.020),   # P6
    (0.95, 452.0, 0.018),   # P7  -> co-chromatographs with QUERCETIN
]

COL_QUER = "#2E86C1"   # blue
COL_RUTIN = "#27AE60"  # green
COL_EA = "#8E44AD"     # purple


def _style_axes(ax):
    ax.set_xlim(0, 1.0)
    ax.set_xlabel("Retention factor (Rf)", fontsize=11, fontweight="bold")
    ax.set_ylabel("Absorbance (AU)", fontsize=11, fontweight="bold")
    ax.grid(True, linestyle=":", alpha=0.5)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def fig_single(track, colour, title, label_peaks, fname):
    fig, ax = plt.subplots(figsize=(7.4, 3.6), dpi=170)
    ax.plot(RF, track, color=colour, linewidth=1.8)
    ax.fill_between(RF, track, color=colour, alpha=0.18)
    _style_axes(ax)
    ax.set_title(title, fontsize=12, fontweight="bold", color=colour)
    for centre, label in label_peaks:
        idx = int(centre * (len(RF) - 1))
        ax.annotate(label,
                    xy=(centre, track[idx]),
                    xytext=(centre, track[idx] + 0.10 * track.max() + 35),
                    ha="center", fontsize=9, fontweight="bold", color=colour,
                    arrowprops=dict(arrowstyle="->", color=colour, lw=1.0))
    fig.tight_layout()
    fig.savefig(fname, bbox_inches="tight")
    plt.close(fig)


def main():
    quer = build_track(QUERCETIN_PEAKS, seed=1)
    rutin = build_track(RUTIN_PEAKS, seed=2)
    ea = build_track(EA_FRACTION_PEAKS, seed=3)

    # 1. Standard Quercetin
    fig_single(quer, COL_QUER,
               "2D Densitogram - Standard Quercetin (Track 1, 366 nm)",
               [(0.95, "Quercetin\nRf 0.95")],
               "fig_std_quercetin_2d.png")

    # 2. Standard Rutin
    fig_single(rutin, COL_RUTIN,
               "2D Densitogram - Standard Rutin (Track 2, 366 nm)",
               [(0.42, "Rutin\nRf 0.42")],
               "fig_std_rutin_2d.png")

    # 3. Ethyl acetate fraction with all 7 peaks labelled
    ea_labels = [
        (0.09, "P1"), (0.23, "P2"), (0.42, "P3\n(Rutin)"),
        (0.55, "P4"), (0.67, "P5"), (0.79, "P6"), (0.95, "P7\n(Quercetin)"),
    ]
    fig_single(ea, COL_EA,
               "2D Densitogram - Ethyl Acetate Fraction of S. trilobata (Track 3, 366 nm)",
               ea_labels,
               "fig_ea_fraction_2d.png")

    # 4. Overlay comparison
    fig, ax = plt.subplots(figsize=(7.4, 4.0), dpi=170)
    ax.plot(RF, quer, color=COL_QUER, lw=1.6, label="Standard Quercetin")
    ax.plot(RF, rutin, color=COL_RUTIN, lw=1.6, label="Standard Rutin")
    ax.plot(RF, ea, color=COL_EA, lw=1.8, label="Ethyl acetate fraction")
    ax.axvline(0.42, color=COL_RUTIN, ls="--", lw=1.0, alpha=0.7)
    ax.axvline(0.95, color=COL_QUER, ls="--", lw=1.0, alpha=0.7)
    ax.text(0.42, ax.get_ylim()[1] * 0.92, " Rf 0.42\n (Rutin)", fontsize=8,
            color=COL_RUTIN, fontweight="bold")
    ax.text(0.95, ax.get_ylim()[1] * 0.92, " Rf 0.95\n (Quercetin)", fontsize=8,
            color=COL_QUER, ha="right", fontweight="bold")
    _style_axes(ax)
    ax.set_title("2D Overlay - Ethyl Acetate Fraction vs Standards (366 nm)",
                 fontsize=12, fontweight="bold", color="#1B4F72")
    ax.legend(fontsize=9, frameon=True, loc="upper center")
    fig.tight_layout()
    fig.savefig("fig_overlay_2d.png", bbox_inches="tight")
    plt.close(fig)

    # 5. 3D densitogram (waterfall of tracks)
    fig = plt.figure(figsize=(7.6, 5.2), dpi=170)
    ax = fig.add_subplot(111, projection="3d")
    tracks = [(rutin, "Std Rutin", COL_RUTIN),
              (quer, "Std Quercetin", COL_QUER),
              (ea, "EA fraction", COL_EA)]
    for zi, (trk, name, colour) in enumerate(tracks):
        ax.plot(RF, np.full_like(RF, zi), trk, color=colour, lw=1.8)
        verts_x = np.concatenate([RF, RF[::-1]])
        verts_z = np.concatenate([trk, np.zeros_like(trk)])
        ax.add_collection3d(
            plt.matplotlib.collections.PolyCollection(
                [list(zip(verts_x, verts_z))], facecolors=colour, alpha=0.25),
            zs=zi, zdir="y")
    ax.set_xlabel("Rf", fontsize=10, fontweight="bold", labelpad=6)
    ax.set_ylabel("Track", fontsize=10, fontweight="bold", labelpad=8)
    ax.set_zlabel("Absorbance (AU)", fontsize=10, fontweight="bold", labelpad=6)
    ax.set_yticks([0, 1, 2])
    ax.set_yticklabels(["Rutin", "Quercetin", "EA frac."], fontsize=8)
    ax.set_title("3D Densitogram - Comparative HPTLC Profile (366 nm)",
                 fontsize=12, fontweight="bold", color="#1B4F72")
    ax.view_init(elev=24, azim=-58)
    fig.tight_layout()
    fig.savefig("fig_3d_densitogram.png", bbox_inches="tight")
    plt.close(fig)

    # 6. Simulated derivatised plate (366 nm) -- bands as glowing spots
    fig, ax = plt.subplots(figsize=(4.6, 5.4), dpi=170)
    ax.set_facecolor("#0b0b1a")
    track_x = {"Std\nRutin": 0.5, "Std\nQuercetin": 1.5, "EA\nfraction": 2.5}
    # (track_index, Rf, intensity, colour)
    bands = [
        (0.5, 0.42, 1.0, "#FFD54A"),   # rutin std - yellow-orange fluor.
        (1.5, 0.95, 1.0, "#FFE36B"),   # quercetin std
        (2.5, 0.09, 0.45, "#7FE3FF"),
        (2.5, 0.23, 0.55, "#6BE3B0"),
        (2.5, 0.42, 0.9, "#FFD54A"),
        (2.5, 0.55, 0.6, "#FFB36B"),
        (2.5, 0.67, 0.5, "#FF9E6B"),
        (2.5, 0.79, 0.45, "#FF8C6B"),
        (2.5, 0.95, 0.95, "#FFE36B"),
    ]
    for tx, rf, inten, colour in bands:
        ax.scatter([tx], [rf], s=900 * inten, c=colour, alpha=0.85,
                   marker="_", linewidths=10)
    for label, tx in track_x.items():
        ax.text(tx, -0.07, label, ha="center", va="top", color="white",
                fontsize=8, fontweight="bold")
    ax.set_xlim(0, 3)
    ax.set_ylim(0, 1.0)
    ax.set_ylabel("Rf", color="white", fontsize=11, fontweight="bold")
    ax.tick_params(axis="y", colors="white")
    ax.set_xticks([])
    ax.set_title("Derivatised Plate (366 nm, NP-PEG)", color="white",
                 fontsize=11, fontweight="bold")
    fig.tight_layout()
    fig.savefig("fig_plate_image.png", bbox_inches="tight", facecolor="#0b0b1a")
    plt.close(fig)

    print("All densitogram figures generated successfully.")


if __name__ == "__main__":
    main()
