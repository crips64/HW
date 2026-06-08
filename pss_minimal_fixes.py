"""
Minimal drop-in fixes for the PSS-DTOF notebook.
Insert this cell after tau_map_all/T_minus/T_plus/delta_Vp are available,
replacing const_matrix, const_y, and compute_p_x_formula_36.
"""

import numpy as np
import pandas as pd


def read_pressure_xyz_xy(path, shape=None):
    """Read tNavigator XYZ with columns X Y PRES into a 2D numpy array."""
    df = pd.read_csv(
        path,
        sep=r"\s+",
        header=None,
        skiprows=5,
        names=["x", "y", "pressure"],
    )
    pmap = (
        df.pivot(index="y", columns="x", values="pressure")
        .sort_index(axis=0)
        .sort_index(axis=1)
        .to_numpy(dtype=float)
    )
    if shape is not None and pmap.shape != tuple(shape):
        raise ValueError(f"pressure map shape={pmap.shape}, expected={shape}")
    return pmap


def pv_average(field, pore_volume):
    """Pore-volume weighted average."""
    field = np.asarray(field, dtype=float)
    pore_volume = np.asarray(pore_volume, dtype=float)
    return float(np.sum(field * pore_volume) / np.sum(pore_volume))


def build_superposition_system_minimal(
    M,
    ct,
    V_res,
    dt,
    p_avg_prev,
    wells,
    controls,
    p_init_map=None,
):
    """
    Unknown vector:
        x = [p_avg, pwf_1, ..., pwf_n, q_1, ..., q_n]

    controls is a list of dictionaries, one per well:
        {"mode": "rate", "value": q_target}
        {"mode": "bhp",  "value": bhp_target}

    If p_init_map is supplied, the well equations include the hydrostatic offset:
        p_wf_i = p_avg + (p0_i - p0_avg) + sum_j M_ij q_j
    """
    M = np.asarray(M, dtype=float)
    n = M.shape[0]

    if M.shape != (n, n):
        raise ValueError(f"M must be square, got {M.shape}")
    if len(wells) != n:
        raise ValueError(f"len(wells)={len(wells)} but M is {M.shape}")
    if len(controls) != n:
        raise ValueError(f"len(controls)={len(controls)} but M is {M.shape}")

    A = np.zeros((2 * n + 1, 2 * n + 1), dtype=float)
    y = np.zeros(2 * n + 1, dtype=float)

    # 1) Material balance for average pressure.
    A[0, 0] = ct * V_res
    A[0, 1 + n : 1 + 2 * n] = dt
    y[0] = ct * V_res * p_avg_prev

    # 2) Well pressure equations.
    for i, well in enumerate(wells):
        row = 1 + i
        A[row, 0] = 1.0
        A[row, 1 + i] = -1.0
        A[row, 1 + n : 1 + 2 * n] = M[i, :]

        if p_init_map is None:
            hydro_offset = 0.0
        else:
            hydro_offset = float(p_init_map[well["rc"]] - p_avg_prev)

        # p_avg - pwf_i + sum_j M_ij q_j = -(p0_i - p0_avg)
        y[row] = -hydro_offset

    # 3) Rate/BHP controls.
    for i, control in enumerate(controls):
        row = 1 + n + i
        mode = control["mode"].lower()
        value = float(control["value"])

        if mode == "rate":
            A[row, 1 + n + i] = 1.0
            y[row] = value
        elif mode == "bhp":
            A[row, 1 + i] = 1.0
            y[row] = value
        else:
            raise ValueError(f"Unknown control mode: {control['mode']}")

    return A, y


def tau_index_map(tau_map):
    tau_map = np.asarray(tau_map)
    tau_values = np.sort(np.unique(tau_map))
    tau_idx = np.searchsorted(tau_values, tau_map)
    return tau_idx, tau_values


def source_pressure_response(
    tau_map,
    time,
    T_minus_j,
    T_plus_j,
    mu,
    delta_Vp_j,
    calculate_inverted_productivity,
):
    tau_map = np.asarray(tau_map)
    T_minus_j = np.asarray(T_minus_j, dtype=float)
    T_plus_j = np.asarray(T_plus_j, dtype=float)
    delta_Vp_j = np.asarray(delta_Vp_j, dtype=float)

    tau_idx, tau_values = tau_index_map(tau_map)
    n_tau = len(tau_values)

    for name, arr in [
        ("T_minus_j", T_minus_j),
        ("T_plus_j", T_plus_j),
        ("delta_Vp_j", delta_Vp_j),
    ]:
        if len(arr) != n_tau:
            raise ValueError(f"{name} length={len(arr)} but n_tau={n_tau}")

    A, J_inv = calculate_inverted_productivity(
        tau_map=tau_map,
        time=time,
        T_minus=T_minus_j,
        T_plus=T_plus_j,
        mu=mu,
        delta_Vp=delta_Vp_j,
    )
    response = A[tau_idx] - J_inv
    return response, A, J_inv, tau_idx, tau_values


def compute_p_x_formula_36_minimal(
    tau_map_all,
    T_minus,
    T_plus,
    delta_Vp,
    time,
    mu,
    ct,
    V_res,
    q_rates,
    p_init,
    calculate_inverted_productivity,
):
    """
    Absolute pressure field:
        p(x,t) = p0(x) + sum_j q_j * [A_j(tau_j(x),t) - J_inv_j(t)]
                 - sum_j q_j * t / (ct * V_res)

    p_init can be scalar or 2D map. q_rates are positive for production
    in the sign convention used by the notebook.
    """
    n_wells = len(tau_map_all)
    q_rates = np.asarray(q_rates, dtype=float).ravel()
    if len(q_rates) != n_wells:
        raise ValueError(f"len(q_rates)={len(q_rates)} but n_wells={n_wells}")

    shape = np.asarray(tau_map_all[0]).shape
    if np.isscalar(p_init):
        p_x = np.full(shape, float(p_init), dtype=float)
    else:
        p_x = np.asarray(p_init, dtype=float).copy()
        if p_x.shape != shape:
            raise ValueError(f"p_init.shape={p_x.shape}, expected={shape}")

    pressure_response_fields = []
    pressure_drop_components = []
    diagnostics = []

    for j in range(n_wells):
        response, A, J_inv, tau_idx, tau_values = source_pressure_response(
            tau_map=tau_map_all[j],
            time=time,
            T_minus_j=T_minus[j],
            T_plus_j=T_plus[j],
            mu=mu,
            delta_Vp_j=delta_Vp[j],
            calculate_inverted_productivity=calculate_inverted_productivity,
        )
        component = q_rates[j] * response
        p_x += component
        pressure_response_fields.append(response)
        pressure_drop_components.append(component)
        diagnostics.append({
            "A": A,
            "J_inv": J_inv,
            "tau_idx": tau_idx,
            "tau_values": tau_values,
        })

    storage_drop = float(np.sum(q_rates) * time / (ct * V_res))
    p_x -= storage_drop

    return p_x, pressure_response_fields, pressure_drop_components, storage_drop, diagnostics
