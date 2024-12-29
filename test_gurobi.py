import gurobipy as gp
from gurobipy import GRB

import pandas as pd
import ast

def write_to_excel(file_path):
    # Define Index Sets
    I = [1, 2]  # Storage Tank 1 and 2
    J = [1, 2]  # Charging Tank 1 and 2
    K = [1]  # Component 1 (could represent a specific crude type or quality)
    L = [1]  # CDU 1
    T = list(range(1, 9))  # Time periods 1 through 8
    V = [1, 2]  # Vessel 1 and 2

    # Define Scalar Parameters
    SCH = 8  # Scheduling Horizon (total number of time periods)
    NCDU = 1  # Number of Crude Distillation Units (CDUs)
    MODE = 1  # Minimum changeover number for CDUs (could relate to operational constraints)

    # Define Single-Indexed Parameters
    CUNLOAD = {
        1: 8.0,  # Vessel 1
        2: 8.0  # Vessel 2
    }
    CSEA = {
        1: 5.0,  # Vessel 1
        2: 5.0  # Vessel 2
    }
    CINVST = {
        1: 0.05,  # Storage Tank 1
        2: 0.05  # Storage Tank 2
    }
    CINVBL = {
        1: 0.08,  # Charging Tank 1
        2: 0.08  # Charging Tank 2
    }
    DM = {
        1: 100.0,  # Blend 1
        2: 100.0  # Blend 2
    }
    FVSMIN = {
        1: 0.0,  # Vessel 1 to any Storage Tank
        2: 0.0  # Vessel 2 to any Storage Tank
    }
    FVSMAX = {
        1: 50.0,  # Vessel 1 to any Storage Tank
        2: 50.0  # Vessel 2 to any Storage Tank
    }
    FSBMIN = {
        1: 0.0,  # Storage Tank 1 to any Charging Tank
        2: 0.0  # Storage Tank 2 to any Charging Tank
    }
    FSBMAX = {
        1: 100.0,  # Storage Tank 1 to any Charging Tank
        2: 100.0  # Storage Tank 2 to any Charging Tank
    }
    FBCMIN = {
        1: 10.0,  # Charging Tank 1 to CDU
        2: 10.0  # Charging Tank 2 to CDU
    }
    FBCMAX = {
        1: 50.0,  # Charging Tank 1 to CDU
        2: 50.0  # Charging Tank 2 to CDU
    }
    TARR = {
        1: 1,  # Vessel 1 arrives at time period 1
        2: 5  # Vessel 2 arrives at time period 5
    }
    TLEA = {
        1: 6,  # Vessel 1 must leave by time period 6
        2: 8  # Vessel 2 must leave by time period 8
    }
    VVESINI = {
        1: 100.0,  # Vessel 1 starts with 100 units
        2: 100.0  # Vessel 2 starts with 100 units
    }
    VSTOINI = {
        1: 25.0,  # Storage Tank 1 starts with 25 units
        2: 75.0  # Storage Tank 2 starts with 75 units
    }
    VSTOMIN = {
        1: 0.0,  # Storage Tank 1 can be empty
        2: 0.0  # Storage Tank 2 can be empty
    }
    VSTOMAX = {
        1: 100.0,  # Storage Tank 1 can hold up to 100 units
        2: 100.0  # Storage Tank 2 can hold up to 100 units
    }
    VBLEINI = {
        1: 50.0,  # Charging Tank 1 starts with 50 units
        2: 50.0  # Charging Tank 2 starts with 50 units
    }
    VBLEMIN = {
        1: 0.0,  # Charging Tank 1 can have 0 units
        2: 0.0  # Charging Tank 2 can have 0 units
    }
    VBLEMAX = {
        1: 100.0,  # Charging Tank 1 can hold up to 100 units
        2: 100.0  # Charging Tank 2 can hold up to 100 units
    }
    DURATION = {
        1: 2,  # Vessel 1 requires at least 2 time periods to unload
        2: 2  # Vessel 2 requires at least 2 time periods to unload
    }
    CSETUP = {
        1: 50  # CDU 1 incurs a setup cost of 50 units
    }

    # Define Multi-Indexed Parameters (Tables)
    VESSTO = {
        1: {1: 1, 2: 0},  # Vessel 1 can transfer to Storage Tank 1 (1) but not to Storage Tank 2 (0)
        2: {1: 0, 2: 1}  # Vessel 2 can transfer to Storage Tank 2 (1) but not to Storage Tank 1 (0)
    }
    STOBLE = {
        1: {1: 1, 2: 1},  # Storage Tank 1 can transfer to both Charging Tanks 1 and 2
        2: {1: 1, 2: 1}  # Storage Tank 2 can transfer to both Charging Tanks 1 and 2
    }
    BLECDU = {
        1: {1: 1},  # Charging Tank 1 can charge CDU 1
        2: {1: 1}  # Charging Tank 2 can charge CDU 1
    }
    VCOMINI = {
        1: {1: 0.1},  # Vessel 1 has 10% concentration of Component 1
        2: {1: 0.6}  # Vessel 2 has 60% concentration of Component 1
    }
    SCOMINI = {
        1: {1: 0.1},  # Storage Tank 1 starts with 10% concentration of Component 1
        2: {1: 0.6}  # Storage Tank 2 starts with 60% concentration of Component 1
    }
    BCOMINI = {
        1: {1: 0.2},  # Charging Tank 1 starts with 20% concentration of Component 1
        2: {1: 0.5}  # Charging Tank 2 starts with 50% concentration of Component 1
    }
    BCOMMIN = {
        1: {1: 0.15},  # Charging Tank 1 must maintain at least 15% concentration of Component 1
        2: {1: 0.45}  # Charging Tank 2 must maintain at least 45% concentration of Component 1
    }
    BCOMMAX = {
        1: {1: 0.25},  # Charging Tank 1 must not exceed 25% concentration of Component 1
        2: {1: 0.55}  # Charging Tank 2 must not exceed 55% concentration of Component 1
    }

    # Combine all data into a dictionary
    data = {
        "Parameter": ["I", "J", "K", "L", "T", "V", "SCH", "NCDU", "MODE", "CUNLOAD", "CSEA", "CINVST", "CINVBL", "DM",
                      "FVSMIN", "FVSMAX", "FSBMIN", "FSBMAX", "FBCMIN", "FBCMAX", "TARR", "TLEA", "VVESINI", "VSTOINI",
                      "VSTOMIN", "VSTOMAX", "VBLEINI", "VBLEMIN", "VBLEMAX", "DURATION", "CSETUP", "VESSTO", "STOBLE",
                      "BLECDU", "VCOMINI", "SCOMINI", "BCOMINI", "BCOMMIN", "BCOMMAX"],
        "Value": [str(I), str(J), str(K), str(L), str(T), str(V), SCH, NCDU, MODE, str(CUNLOAD), str(CSEA),
                  str(CINVST), str(CINVBL), str(DM), str(FVSMIN), str(FVSMAX), str(FSBMIN), str(FSBMAX), str(FBCMIN),
                  str(FBCMAX), str(TARR), str(TLEA), str(VVESINI), str(VSTOINI), str(VSTOMIN), str(VSTOMAX),
                  str(VBLEINI), str(VBLEMIN), str(VBLEMAX), str(DURATION), str(CSETUP), str(VESSTO), str(STOBLE),
                  str(BLECDU), str(VCOMINI), str(SCOMINI), str(BCOMINI), str(BCOMMIN), str(BCOMMAX)]
    }

    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)


def read_from_excel(file_path):
    df = pd.read_excel(file_path)
    # Convert dictionaries stored as strings back to actual dictionaries
    def str_to_dict(s):
        try:
            return ast.literal_eval(s)
        except (ValueError, SyntaxError):
            return s

    df['Value'] = df['Value'].apply(str_to_dict)
    result_dict = df.set_index('Parameter')['Value'].to_dict()
    # Reassign variables from the dictionary
    I = result_dict["I"]
    J = result_dict["J"]
    K = result_dict["K"]
    L = result_dict["L"]
    T = result_dict["T"]
    V = result_dict["V"]
    SCH = result_dict["SCH"]
    NCDU = result_dict["NCDU"]
    MODE = result_dict["MODE"]
    CUNLOAD = result_dict["CUNLOAD"]
    CSEA = result_dict["CSEA"]
    CINVST = result_dict["CINVST"]
    CINVBL = result_dict["CINVBL"]
    DM = result_dict["DM"]
    FVSMIN = result_dict["FVSMIN"]
    FVSMAX = result_dict["FVSMAX"]
    FSBMIN = result_dict["FSBMIN"]
    FSBMAX = result_dict["FSBMAX"]
    FBCMIN = result_dict["FBCMIN"]
    FBCMAX = result_dict["FBCMAX"]
    TARR = result_dict["TARR"]
    TLEA = result_dict["TLEA"]
    VVESINI = result_dict["VVESINI"]
    VSTOINI = result_dict["VSTOINI"]
    VSTOMIN = result_dict["VSTOMIN"]
    VSTOMAX = result_dict["VSTOMAX"]
    VBLEINI = result_dict["VBLEINI"]
    VBLEMIN = result_dict["VBLEMIN"]
    VBLEMAX = result_dict["VBLEMAX"]
    DURATION = result_dict["DURATION"]
    CSETUP = result_dict["CSETUP"]
    VESSTO = result_dict["VESSTO"]
    STOBLE = result_dict["STOBLE"]
    BLECDU = result_dict["BLECDU"]
    VCOMINI = result_dict["VCOMINI"]
    SCOMINI = result_dict["SCOMINI"]
    BCOMINI = result_dict["BCOMINI"]
    BCOMMIN = result_dict["BCOMMIN"]
    BCOMMAX = result_dict["BCOMMAX"]
    return (I, J, K, L, T, V, SCH, NCDU, MODE, CUNLOAD, CSEA, CINVST, CINVBL, DM, FVSMIN, FVSMAX, FSBMIN, FSBMAX, FBCMIN,
            FBCMAX, TARR, TLEA, VVESINI, VSTOINI, VSTOMIN, VSTOMAX, VBLEINI, VBLEMIN, VBLEMAX, DURATION, CSETUP, VESSTO,
            STOBLE, BLECDU, VCOMINI, SCOMINI, BCOMINI, BCOMMIN, BCOMMAX)


import pandas as pd


def print_and_save_to_excel(model, V, I, J, L, T, X_W_vars, X_F_vars, X_L_vars, F_VS_vars, F_BS_vars, F_BC_vars,
                            V_S_vars, V_B_vars, excel_path):
    # Create a dictionary to store the data
    data = {}
    # Store the status of the solution
    data["Status"] = model.Status
    # Store the optimal values of decision variables
    for v in V:
        for t in T:
            data[f"X_W[{v}, {t}]"] = X_W_vars[v, t].X
            data[f"X_F[{v}, {t}]"] = X_F_vars[v, t].X
            data[f"X_L[{v}, {t}]"] = X_L_vars[v, t].X
    for v in V:
        for i in I:
            for t in T:
                data[f"F_VS[{v}, {i}, {t}]"] = F_VS_vars[v, i, t].X
    for i in I:
        for j in J:
            for t in T:
                data[f"F_BS[{i}, {j}, {t}]"] = F_BS_vars[i, j, t].X
    for j in J:
        for l in L:
            for t in T:
                data[f"F_BC[{j}, {l}, {t}]"] = F_BC_vars[j, l, t].X
    for j in J:
        for t in T:
            data[f"V_B[{j}, {t}]"] = V_B_vars[j, t].X

    # Special handling for V_S variables to store them in a column-like format
    v_s_column = []
    for i in I:
        for t in T:
            v_s_column.append({
                "Parameter": f"V_S[{i}, {t}]",
                "Value": V_S_vars[i, t].X
            })
    df_v_s = pd.DataFrame(v_s_column)
    data["V_S_entries"] = df_v_s

    # Store the optimal objective value
    data["Total cost"] = model.ObjVal

    # Convert the dictionary to a DataFrame
    df = pd.DataFrame()
    for key, value in data.items():
        if isinstance(value, pd.DataFrame):
            # If the value is a DataFrame, append its rows to the main DataFrame
            df = pd.concat([df, value], ignore_index=True)
        else:
            # If the value is not a DataFrame, create a new entry in the DataFrame
            df = pd.concat([df, pd.DataFrame({"Parameter": [key], "Value": [value]})], ignore_index=True)

    # Save the DataFrame to an Excel file
    df.to_excel(excel_path, index=False)



if __name__ == "__main__":

    file_path = "data.xlsx"  # 这里可以修改为你想要的文件路径
    write_to_excel(file_path)
    (I, J, K, L, T, V, SCH, NCDU, MODE, CUNLOAD, CSEA, CINVST, CINVBL, DM, FVSMIN, FVSMAX, FSBMIN, FSBMAX, FBCMIN, FBCMAX, TARR, TLEA, VVESINI, VSTOINI, VSTOMIN, VSTOMAX, VBLEINI, VBLEMIN, VBLEMAX, DURATION, CSETUP, VESSTO, STOBLE, BLECDU, VCOMINI, SCOMINI, BCOMINI, BCOMMIN, BCOMMAX) = read_from_excel(file_path)


    # ============================================
    # 5. Define the Linear Programming Problem using pygurobi and build QUBO matrix
    # ============================================

    # Create a new model
    model = gp.Model("Crude_Inventory_Management")

    # Define the decision variables
    # X_{W, v, t}: 1 if vessel v unloads in time t, 0 otherwise
    X_W_vars = model.addVars(V, T, vtype=GRB.BINARY, name="X_W")

    # X_{F, v, t}: 1 if vessel v starts unloading in time t, 0 otherwise
    X_F_vars = model.addVars(V, T, vtype=GRB.BINARY, name="X_F")

    # X_{L, v, t}: 1 if vessel v leaves after unloading in time t, 0 otherwise
    X_L_vars = model.addVars(V, T, vtype=GRB.BINARY, name="X_L")

    # F_{VS, v, i, t}: crude oil transfer rate from vessel v to storage tank i in time t
    F_VS_vars = model.addVars(V, I, T, lb=0, ub=max(FVSMAX.values()), vtype=GRB.CONTINUOUS, name="F_VS")

    # F_{BS, i, j, t}: crude oil transfer rate from storage tank i to charging tank j in time t
    F_BS_vars = model.addVars(I, J, T, lb=0, ub=max(FSBMAX.values()), vtype=GRB.CONTINUOUS, name="F_BS")

    # F_{BC, j, l, t}: charging rate from charging tank j to CDU l in time t
    F_BC_vars = model.addVars(J, L, T, lb=0, ub=max(FBCMAX.values()), vtype=GRB.CONTINUOUS, name="F_BC")

    # V_{S, i, t}: volume of crude oil in storage tank i in time t
    V_S_vars = model.addVars(I, T, lb=0, ub=max(VSTOMAX.values()), vtype=GRB.CONTINUOUS, name="V_S")

    # V_{B, j, t}: volume of mixed crude oil in charging tank j in time t
    V_B_vars = model.addVars(J, T, lb=0, ub=max(VBLEMAX.values()), vtype=GRB.CONTINUOUS, name="V_B")

    # Z_{j, j', l, t}: 1 if switching from blend j to blend j' happens in CDU l at time t, 0 otherwise
    Z_vars = model.addVars(J, J, L, T, vtype=GRB.BINARY, name="Z", ub=1, lb=0)

    # Objective function terms
    obj_terms = []
    # Unloading cost term
    for v in V:
        unloading_cost_term = CUNLOAD[v] * gp.quicksum((X_W_vars[v, t] - X_W_vars[v, t - 1]) for t in T[1:])
        obj_terms.append(unloading_cost_term)

    # Sea waiting cost term
    for v in V:
        sea_waiting_cost_term = CSEA[v] * gp.quicksum(
            (X_F_vars[v, t] - X_F_vars[v, t - 1]) * (t - TARR[v]) for t in T[1:])
        obj_terms.append(sea_waiting_cost_term)

    # Storage tank inventory cost term
    for i in I:
        storage_inventory_cost_term = CINVST[i] * gp.quicksum((V_S_vars[i, t] + V_S_vars[i, t - 1]) / 2 for t in T[1:])
        obj_terms.append(storage_inventory_cost_term)

    # Charging tank inventory cost term
    for j in J:
        charging_inventory_cost_term = CINVBL[j] * gp.quicksum((V_B_vars[j, t] + V_B_vars[j, t - 1]) / 2 for t in T[1:])
        obj_terms.append(charging_inventory_cost_term)

    # Setup cost term
    for l in L:
        setup_cost_term = CSETUP[l] * gp.quicksum(
            Z_vars[j, j_prime, l, t] for j in J for j_prime in J if j != j_prime for t in T)
        obj_terms.append(setup_cost_term)

    # Set the objective function
    model.setObjective(gp.quicksum(obj_terms), GRB.MINIMIZE)

    # Constraints

    # (1) Each vessel arrives only once and leaves only once
    for v in V:
        model.addConstr(gp.quicksum(X_F_vars[v, t] for t in T) == 1, name=f"Arrival_{v}")
        model.addConstr(gp.quicksum(X_L_vars[v, t] for t in T) == 1, name=f"Departure_{v}")

    # (2) If X_{F, v, t} = 1, then X_{W, v, t+1} = 1,..., X_{W, v, t+DURATION[v]} = 1
    for v in V:
        for t in range(1, max(T) + 1 - DURATION[v] + 1):
            model.addConstr(
                gp.quicksum(X_W_vars[v, t_prime] for t_prime in range(t, t + DURATION[v])) >= DURATION[v] * X_F_vars[v,
                t],
                name=f"Unloading_Start_{v}_{t}")

    # (3) If X_{F, v, t} = 1, then t >= T_{ARR, v}
    for v in V:
        for t in T:
            if t < TARR[v]:
                model.addConstr(X_F_vars[v, t] == 0, name=f"Unloading_Time_Before_Arrival_{v}_{t}")
            else:
                model.addConstr(X_F_vars[v, t] <= X_W_vars[v, t], name=f"Unloading_Time_{v}_{t}")

    # (4) If X_{L, v, t} = 1, then t <= T_{LEA, v}
    for v in V:
        for t in T:
            if t > TLEA[v]:
                model.addConstr(X_L_vars[v, t] == 0, name=f"Leaving_Time_After_Leave_{v}_{t}")
            else:
                model.addConstr(X_L_vars[v, t] >= X_W_vars[v, t], name=f"Leaving_Time_{v}_{t}")

    # (5) Only one vessel can unload at a time
    for t in T:
        model.addConstr(gp.quicksum(X_F_vars[v, t] for v in V) <= 1, name=f"Unload_Unique_{t}")

    # (6) Material balance equation for vessels
    for v in V:
        model.addConstr(VVESINI[v] == gp.quicksum(F_VS_vars[v, i, t] for i in I for t in T),
                        name=f"VES_Oil_Balance_{v}")

    # (7) Material balance equation for storage tanks
    for i in I:
        model.addConstr(V_S_vars[i, 1] == VSTOINI[i], name=f"Initial_Inventory_Storage_{i}")
        for t in range(2, len(T) + 1):
            model.addConstr(V_S_vars[i, t] == V_S_vars[i, t - 1] + gp.quicksum(F_VS_vars[v, i, t] for v in V) -
                            gp.quicksum(F_BS_vars[i, j, t] for j in J), name=f"STO_Oil_Balance_{i}_{t}")

    # (8) Material balance equation for charging tanks
    for j in J:
        model.addConstr(V_B_vars[j, 1] == VBLEINI[j], name=f"Initial_Inventory_Charging_{j}")
        for t in range(2, len(T) + 1):
            model.addConstr(V_B_vars[j, t] == V_B_vars[j, t - 1] + gp.quicksum(F_BS_vars[i, j, t] for i in I) -
                            gp.quicksum(F_BC_vars[j, l, t] for l in L), name=f"BLE_Oil_Balance_{j}_{t}")

    # (9) Demand for each blend by CDUs
    for j in J:
        model.addConstr(gp.quicksum(F_BC_vars[j, l, t] for l in L for t in T) == DM[j], name=f"Blend_Demand_{j}")

    # (10) Volume constraints for storage tanks
    for i in I:
        for t in T:
            model.addConstr(V_S_vars[i, t] <= VSTOMAX[i], name=f"Max_Inventory_Storage_{i}_{t}")
            model.addConstr(V_S_vars[i, t] >= VSTOMIN[i], name=f"Min_Inventory_Storage_{i}_{t}")

    # (11) Volume constraints for charging tanks
    for j in J:
        for t in T:
            model.addConstr(V_B_vars[j, t] <= VBLEMAX[j], name=f"Max_Inventory_Charging_{j}_{t}")
            model.addConstr(V_B_vars[j, t] >= VBLEMIN[j], name=f"Min_Inventory_Charging_{j}_{t}")

    # (12) Component concentration constraints in charging tanks
    for j in J:
        for k in K:
            for t in T:
                left_side_expr = (
                            BCOMINI[j][k] * VBLEINI[j] + gp.quicksum(F_BS_vars[i, j, t] * SCOMINI[i][k] for i in I)
                            - gp.quicksum(F_BC_vars[j, l, t] * BCOMINI[j][k] for l in L))
                model.addConstr(left_side_expr >= V_B_vars[j, t] * BCOMMIN[j][k],
                                name=f"Concentration_Min_J{j}_K{k}_T{t}")
                model.addConstr(left_side_expr <= V_B_vars[j, t] * BCOMMAX[j][k],
                                name=f"Concentration_Max_J{j}_K{k}_T{t}")

    # (13) Transfer rate constraints for vessels to storage tanks
    for v in V:
        for i in I:
            for t in T:
                model.addConstr(F_VS_vars[v, i, t] <= VVESINI[v] * X_F_vars[v, t], name=f"Max_F_VS_{v}_{i}_{t}_X_F")
                model.addConstr(F_VS_vars[v, i, t] >= FVSMIN[v] * X_F_vars[v, t], name=f"Min_F_VS_{v}_{i}_{t}")
                model.addConstr(F_VS_vars[v, i, t] <= FVSMAX[v] * X_W_vars[v, t], name=f"Max_F_VS_{v}_{i}_{t}_X_W")

    # (14) Transfer rate constraints for storage tanks to charging tanks
    for i in I:
        for j in J:
            for t in T:
                model.addConstr(F_BS_vars[i, j, t] <= FSBMAX[i], name=f"Max_F_BS_{i}_{j}_{t}")
                model.addConstr(F_BS_vars[i, j, t] >= FSBMIN[i], name=f"Min_F_BS_{i}_{j}_{t}")

    # (15) Charging rate constraints from charging tanks to CDUs
    for j in J:
        for l in L:
            for t in T:
                model.addConstr(F_BC_vars[j, l, t] <= FBCMAX[j], name=f"Max_F_BC_{j}_{l}_{t}")
                model.addConstr(F_BC_vars[j, l, t] >= FBCMIN[j], name=f"Min_F_BC_{j}_{l}_{t}")

    # (16) Setup constraints for CDUs
    for l in L:
        for j in J:
            for j_prime in J:
                if j != j_prime:
                    for t in T:
                        model.addConstr(
                            gp.quicksum(F_BC_vars[j_prime, l, t_prime] for t_prime in range(1, t + 1)) >= MODE * Z_vars[
                                j, j_prime, l, t], name=f"Setup_Constraint_L{l}_J{j}_J{j_prime}_T{t}")
                        model.addConstr(
                            gp.quicksum(F_BC_vars[j, l, t_prime] for t_prime in range(1, t + 1)) >= MODE * Z_vars[
                                j_prime, j, l, t], name=f"Setup_Constraint_Rev_L{l}_J{j_prime}_J{j}_T{t}")
                        # Ensure that Z is 1 if there is a switch
                        model.addConstr(
                            F_BC_vars[j, l, t] + F_BC_vars[j_prime, l, t] <= (VBLEMAX[j] + VBLEMAX[j_prime]) * (
                                        1 - Z_vars[j, j_prime, l, t]),
                            name=f"Switch_Constraint_L{l}_J{j}_J{j_prime}_T{t}")
                        model.addConstr(
                            F_BC_vars[j_prime, l, t] + F_BC_vars[j, l, t] <= (VBLEMAX[j] + VBLEMAX[j_prime]) * (
                                        1 - Z_vars[j_prime, j, l, t]),
                            name=f"Switch_Constraint_Rev_L{l}_J{j_prime}_J{j}_T{t}")

    # (17) No blending between different components (simplified constraint)
    for j in J:
        for k in K:
            for t in T:
                left_side_expr = (
                            BCOMINI[j][k] * VBLEINI[j] + gp.quicksum(F_BS_vars[i, j, t] * SCOMINI[i][k] for i in I))
                model.addConstr(left_side_expr <= VBLEMAX[j] * BCOMMAX[j][k], name=f"Blending_Max_J{j}_K{k}_T{t}")

    # Solve the model
    model.optimize()

    # Print the status of the solution
    print(f"Status: {model.Status}")

    # Print the optimal values of the decision variables
    print("Optimal values of decision variables:")
    for v in V:
        for t in T:
            print(f"X_W[{v}, {t}] = {X_W_vars[v, t].X}")
            print(f"X_F[{v}, {t}] = {X_F_vars[v, t].X}")
            print(f"X_L[{v}, {t}] = {X_L_vars[v, t].X}")

    for v in V:
        for i in I:
            for t in T:
                print(f"F_VS[{v}, {i}, {t}] = {F_VS_vars[v, i, t].X}")

    for i in I:
        for j in J:
            for t in T:
                print(f"F_BS[{i}, {j}, {t}] = {F_BS_vars[i, j, t].X}")

    for j in J:
        for l in L:
            for t in T:
                print(f"F_BC[{j}, {l}, {t}] = {F_BC_vars[j, l, t].X}")

    for i in I:
        for t in T:
            print(f"V_S[{i}, {t}] = {V_S_vars[i, t].X}")

    for j in J:
        for t in T:
            print(f"V_B[{j}, {t}] = {V_B_vars[j, t].X}")

    # Print the optimal objective value
    print(f"Total cost: {model.ObjVal}")

    # model, V, I, J, L, T, X_W_vars, X_F_vars, X_L_vars, F_VS_vars, F_BS_vars, F_BC_vars, V_S_vars, V_B_vars are defined
    excel_path = "solution.xlsx"
    print_and_save_to_excel(model, V, I, J, L, T, X_W_vars, X_F_vars, X_L_vars, F_VS_vars, F_BS_vars, F_BC_vars,
                            V_S_vars, V_B_vars, excel_path)

    import matplotlib.pyplot as plt

    import matplotlib.pyplot as plt

    # 图 1
    # 创建一个新的 Figure 对象
    plt.figure()
    x = [0, 1, 1, 3, 4, 5, 6, 6, 8]
    y = [0, 0, 100, 0, 0, 0, 0, 100, 0]
    plt.plot(x, y, color='red')
    plt.title('VESSEL UNLOADING SCHEDULE')
    plt.xlabel('time')
    plt.ylabel('volume')
    plt.xlim(0, 8)
    plt.ylim(0, 100)
    plt.savefig('VESSEL_UNLOADING_SCHEDULE.png')

    import matplotlib.pyplot as plt

    # 假设以下变量已经在前面的代码中被赋值
    F_VS_vars = {
        (1, 1, 1): 0.0, (1, 1, 2): 0.0, (1, 1, 3): 0.0, (1, 1, 4): 0.0, (1, 1, 5): 0.0, (1, 1, 6): 0.0, (1, 1, 7): 50.0,
        (1, 1, 8): 0.0,
        (1, 2, 1): 0.0, (1, 2, 2): 0.0, (1, 2, 3): 0.0, (1, 2, 4): 0.0, (1, 2, 5): 0.0, (1, 2, 6): 0.0, (1, 2, 7): 50.0,
        (1, 2, 8): 0.0,
        (2, 1, 1): 0.0, (2, 1, 2): 0.0, (2, 1, 3): 0.0, (2, 1, 4): 0.0, (2, 1, 5): 0.0, (2, 1, 6): 0.0, (2, 1, 7): 0.0,
        (2, 1, 8): 50.0,
        (2, 2, 1): 0.0, (2, 2, 2): 0.0, (2, 2, 3): 0.0, (2, 2, 4): 0.0, (2, 2, 5): 0.0, (2, 2, 6): 0.0, (2, 2, 7): 0.0,
        (2, 2, 8): 50.0
    }

    F_BS_vars = {
        (1, 1, 1): 0.0, (1, 1, 2): 15.278711965778228, (1, 1, 3): 0.0, (1, 1, 4): 0.0, (1, 1, 5): 9.721288034221772,
        (1, 1, 6): 0.0, (1, 1, 7): 18.352667806567602, (1, 1, 8): 11.175559355225335,
        (1, 2, 1): 0.0, (1, 2, 2): 0.0, (1, 2, 3): 0.0, (1, 2, 4): 0.0, (1, 2, 5): 0.0, (1, 2, 6): 0.0,
        (1, 2, 7): 16.262440264807694, (1, 2, 8): 8.608346607820536,
        (2, 1, 1): 5.833333333333333, (2, 1, 2): 10.127734188148212, (2, 1, 3): 6.046035409735175,
        (2, 1, 4): 3.2217749881174416, (2, 1, 5): 2.5464519942963713, (2, 1, 6): 0.0, (2, 1, 7): 0.0, (2, 1, 8): 0.0,
        (2, 2, 1): 12.5, (2, 2, 2): 23.636196238327024, (2, 2, 3): 9.996317243193477, (2, 2, 4): 9.955806918321718,
        (2, 2, 5): 9.469683019860582, (2, 2, 6): 0.0, (2, 2, 7): 0.0, (2, 2, 8): 0.0
    }

    F_BC_vars = {
        (1, 1, 1): 10.0, (1, 1, 2): 24.941996580297193, (1, 1, 3): 10.0, (1, 1, 4): 10.0, (1, 1, 5): 10.0,
        (1, 1, 6): 10.0, (1, 1, 7): 15.058003419702807, (1, 1, 8): 10.0,
        (2, 1, 1): 10.0, (2, 1, 2): 26.363803761673072, (2, 1, 3): 10.0, (2, 1, 4): 10.0, (2, 1, 5): 10.0,
        (2, 1, 6): 13.636196238326928, (2, 1, 7): 10.0, (2, 1, 8): 10.0,
        (2, 2, 1): 0.0, (2, 2, 2): 0.0, (2, 2, 3): 0.0, (2, 2, 4): 0.0, (2, 2, 5): 0.0, (2, 2, 6): 0.0, (2, 2, 7): 0.0,
        (2, 2, 8): 0.0
    }


    def plot_graph(data, var_name, indices, x_label='time', y_label='rate'):
        for index in indices:
            # 创建一个新的 Figure 对象
            plt.figure()
            x = []
            y = []
            # 定义时间范围，这里假设时间范围是 1 到 8
            time_range = range(1, 9)
            for t in time_range:
                key = tuple(index + (t,))
                x.append(t)
                y.append(data.get(key, 0.0))
            plt.plot(x, y, color='red')
            plt.xlim(1, 8)
            plt.ylim(0, 100)
            plt.xlabel(x_label)
            plt.ylabel(y_label)
            plt.title(f'{var_name}[{index[0]}, {index[1]}]')
            plt.savefig(f'{var_name}_{index[0]}_{index[1]}.png')


    # 绘制 F_VS_vars 的图
    v_i_indices = [(v, i) for v in [1, 2] for i in [1, 2]]
    plot_graph(F_VS_vars, 'Transfer rate constraints for vessels to storage tanks', v_i_indices, y_label='volume')

    # 绘制 F_BS_vars 的图
    i_j_indices = [(i, j) for i in [1, 2] for j in [1, 2]]
    plot_graph(F_BS_vars, 'crude oil transfer rate from storage tank i to charging tank j in time t', i_j_indices)

    # 绘制 F_BC_vars 的图
    j_l_indices = [(j, l) for j in [1, 2] for l in [1, 2]]
    plot_graph(F_BC_vars, 'Charging rate constraints from charging tanks to CDUs', j_l_indices)

    import matplotlib.pyplot as plt

    # 存储 V_S_vars 和 V_B_vars 的数据
    V_S_vars = {
        (1, 1): 25.0, (1, 2): 9.721288034221772, (1, 3): 9.721288034221772, (1, 4): 9.721288034221772, (1, 5): 0.0,
        (1, 6): 0.0, (1, 7): 15.384891928624704, (1, 8): 45.60098596557883,
        (2, 1): 75.0, (2, 2): 41.236069573524766, (2, 3): 25.193716920596113, (2, 4): 12.016135014156953, (2, 5): 0.0,
        (2, 6): 0.0, (2, 7): 50.0, (2, 8): 100.0
    }

    V_B_vars = {
        (1, 1): 50.0, (1, 2): 50.464449573629246, (1, 3): 46.51048498336442, (1, 4): 39.73225997148186, (1, 5): 42.0,
        (1, 6): 32.0, (1, 7): 35.294664386864795, (1, 8): 36.470223742090134,
        (2, 1): 50.0, (2, 2): 47.27239247665395, (2, 3): 47.26870971984743, (2, 4): 47.22451663816914,
        (2, 5): 46.694199658029724, (2, 6): 33.058003419702786, (2, 7): 39.32044368451049, (2, 8): 37.928790292331
    }


    def plot_graph(data, var_name, index, title_template, x_label='time', y_label='volume'):
        # 创建一个新的 Figure 对象
        plt.figure()
        x = []
        y = []
        # 假设时间范围是 1 到 8
        time_range = range(1, 9)
        for t in time_range:
            # 对于 V_S_vars 和 V_B_vars，我们使用 index 作为第一个元素，t 作为第二个元素
            key = (index, t)
            x.append(t)
            y.append(data.get(key, 0.0))
        plt.plot(x, y, color='red')
        plt.xlim(1, 8)
        plt.ylim(0, 100)
        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.title(title_template.format(index))
        plt.savefig(f'{var_name}_{index}.png')


    # 绘制 V_S_vars 的图
    plot_graph(V_S_vars, 'V_S', 1, title_template='volume of crude oil in storage tank 1 in time {{}}')
    plot_graph(V_S_vars, 'V_S', 2, title_template='volume of crude oil in storage tank 2 in time {{}}')

    # 绘制 V_B_vars 的图
    plot_graph(V_B_vars, 'V_B', 1, title_template='volume of mixed crude oil in charging tank 1 in time {{}}')
    plot_graph(V_B_vars, 'V_B', 2, title_template='volume of mixed crude oil in charging tank 2 in time {{}}')