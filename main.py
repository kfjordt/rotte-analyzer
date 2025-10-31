import networkx as nx
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook

import itertools

RAT_COUNT = 10


def read_data(path, sheet_name=0):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[sheet_name]
    data = [[cell.value for cell in row] for row in ws.iter_rows()]

    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        value = ws.cell(row=min_row, column=min_col).value
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                data[r - 1][c - 1] = value

    resident_dict = {}
    columns = np.asarray(data).T[1:]
    for column in columns:
        month_label = f"{column[0]} {column[1]}"
        residents = list([resident for resident in column[2:] if resident is not None])
        for resident in residents:
            if resident not in resident_dict:
                resident_dict[resident] = []
            resident_dict[resident].append(month_label)
    return resident_dict

def build_graph(months_by_residents: dict):
    unique_months = sorted(set([m for l in months_by_residents.values() for m in l]))

    residents_by_month = {month: [] for month in unique_months}
    for resident, months in months_by_residents.items():
        for month in months:
            residents_by_month[month].append(resident)

    roomie_graph = {resident: set() for resident in months_by_residents.keys()}
    for month in unique_months:
        residents_in_month = residents_by_month[month]
        for resident in residents_in_month:
            roomie_graph[resident].update(r for r in residents_in_month if r != resident)

    roomie_graph = {r: sorted(list(neighbors)) for r, neighbors in roomie_graph.items()}
    return roomie_graph


def plot_graph(roomie_graph: dict):
    G = nx.Graph()
    for resident, roommates in roomie_graph.items():
        for roommate in roommates:
            G.add_edge(resident, roommate)


    # Total nodes and edges
    num_nodes = G.number_of_nodes()
    num_edges = G.number_of_edges()

    # Compute all shortest paths
    all_paths = dict(nx.all_pairs_shortest_path(G))

    # Flatten into lengths and track longest path
    longest_length = 0
    longest_path = []

    lengths = []
    for u, v in itertools.combinations(G.nodes, 2):
        if v in all_paths[u]:
            path = all_paths[u][v]
            path_len = len(path) - 1  # number of edges
            lengths.append(path_len)
            if path_len > longest_length:
                longest_length = path_len
                longest_path = path

    # Average length
    average_length = sum(lengths) / len(lengths)

    print("Total nodes:", num_nodes)
    print("Total edges:", num_edges)
    print("Longest path length:", longest_length)
    print("Longest path nodes:", longest_path)
    print("Average path length:", average_length)
    pos = nx.kamada_kawai_layout(G)

    # Node properties
    degrees = dict(G.degree())
    node_sizes = [400 + 150 * degrees[n] for n in G.nodes()]
    node_colors = [degrees[n] for n in G.nodes()]

    # A3 portrait page (in inches)
    plt.figure(figsize=(11.7, 16.5), facecolor="white")
    ax = plt.gca()
    ax.set_facecolor("white")

    # Draw edges
    nx.draw_networkx_edges(G, pos, alpha=0.25, width=1.0, edge_color="gray")

    # Draw nodes
    nodes = nx.draw_networkx_nodes(
        G,
        pos,
        node_color=node_colors,
        cmap=plt.cm.Oranges,  # better contrast on white
        node_size=node_sizes,
        edgecolors="black",
        linewidths=0.4,
        alpha=0.95,
    )

    # Labels
    nx.draw_networkx_labels(G, pos, font_size=8, font_weight="bold", font_color="red")

    # Colorbar with black text
    cbar = plt.colorbar(nodes, shrink=0.8, pad=0.02)
    cbar.set_label("Number of Roommates", color="black", fontsize=10)
    plt.setp(plt.getp(cbar.ax.axes, "yticklabels"), color="black")

    # Title and layout
    plt.title("Roommate Network", fontsize=18, fontweight="bold", color="black", pad=20)
    plt.axis("off")
    plt.tight_layout()
    
    # Save to PNG with transparent background
    plt.savefig("graph.png", dpi=300, transparent=True)
    plt.close()


def plot_box_plot(data: list[float]):
    plt.boxplot(data, vert=True, patch_artist=True)
    plt.title("Box Plot")
    plt.tight_layout()
    plt.ylabel("Values")
    plt.show()
    # plt.savefig("boxplot.pdf")


if __name__ == "__main__":
    # prepare data
    months_by_residents = read_data("rottedata.xlsx")

    # for resident, months in sorted(months_by_residents.items()):
    for resident, months in sorted(months_by_residents.items(), key=lambda x: len(x[0]), reverse=True)[:3]:
        print(resident, months)

    # print(months_by_residents)
    roomie_graph = build_graph(months_by_residents)

    # roomie related statistics
    roomie_count_by_resident = {r: len(roomies) for r, roomies in roomie_graph.items()}
    roomie_counts = list(roomie_count_by_resident.values())

    # residency period related statistics
    month_count_by_resident = {r: len(months) for r, months in months_by_residents.items()}
    month_counts = list(month_count_by_resident.values())


    # for resident, count in sorted(roomie_count_by_resident.items(), key=lambda x: x[1], reverse=True)[:3]:
    #     print(f"{resident}: {count}")
    # print(f"Average roomie count: {np.mean(roomie_counts):.1f}")
    # print(f"Median roomie count: {np.median(roomie_counts):.1f}")
    # print(f"Min residency period: {np.min(month_counts):.1f}")
    # print(f"Max residency period: {np.max(month_counts):.1f}")
    # print(f"Average residency period: {np.mean(month_counts):.1f}")
    # print(f"Median residency period: {np.median(month_counts):.1f}")

    # for resident, count in sorted(month_count_by_resident.items(), key=lambda x: x[1], reverse=True)[:3]:
    #     print(f"{resident}: {count}")

    # plot_box_plot(month_counts)
    # plot_graph(roomie_graph)
