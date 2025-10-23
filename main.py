import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import numpy as np

import itertools

RAT_COUNT = 10


def read_data(path: str):
    data = pd.read_excel(path)

    resident_dict = {}

    current_year = None
    current_names = [None] * RAT_COUNT
    for col_name, col_data in list(data.items())[1:]:
        if "Unnamed" not in str(col_name):
            current_year = col_name

        month = col_data[0]
        for name_idx, name in enumerate(col_data[1 : RAT_COUNT + 1]):
            if "nan" not in str(name):
                current_names[name_idx] = name
                
        for name in current_names:
            if name is None or name == " ":
                continue
            if name not in resident_dict:
                resident_dict[name] = []

            resident_dict[name].append(f"{current_year} {month}")

    return {k: list(set(v)) for k, v in resident_dict.items()}


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
    roomie_graph = build_graph(months_by_residents)

    # roomie related statistics
    roomie_count_by_resident = {r: len(roomies) for r, roomies in roomie_graph.items()}
    roomie_counts = list(roomie_count_by_resident.values())

    print(f"Average roomie count: {np.mean(roomie_counts):.1f}")
    print(f"Median roomie count: {np.median(roomie_counts):.1f}")

    for resident, count in sorted(roomie_count_by_resident.items(), key=lambda x: x[1], reverse=True)[:3]:
        print(f"{resident}: {count}")

    # residency period related statistics
    month_count_by_resident = {r: len(months) for r, months in months_by_residents.items()}
    print(sorted(months_by_residents["Leonora Maria Meier-Nielsen"]))
    month_counts = list(month_count_by_resident.values())

    print(f"Min residency period: {np.min(month_counts):.1f}")
    print(f"Max residency period: {np.max(month_counts):.1f}")
    print(f"Average residency period: {np.mean(month_counts):.1f}")
    print(f"Median residency period: {np.median(month_counts):.1f}")

    for resident, count in sorted(month_count_by_resident.items(), key=lambda x: x[1], reverse=True)[:3]:
        print(f"{resident}: {count}")

    # plot_box_plot(month_counts)
    plot_graph(roomie_graph)
