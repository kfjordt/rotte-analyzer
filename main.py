import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt

RAT_COUNT = 9

def read_data(path: str):
    data = pd.read_excel(path)

    resident_dict = {}

    current_year = None
    current_names = [None] * RAT_COUNT
    for col_name, col_data in list(data.items())[1:]:
        if "Unnamed" not in str(col_name):
            current_year = col_name

        month = col_data[0]
        for name_idx, name in enumerate(col_data[1:RAT_COUNT+1]):
            if "nan" not in str(name):
                current_names[name_idx] = name
            
        for name in current_names:
            if name not in resident_dict:
                resident_dict[name] = []
            
            resident_dict[name].append(f"{current_year} {month}")

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

    # Use Kamada–Kawai for even spacing
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
        G, pos,
        node_color=node_colors,
        cmap=plt.cm.cividis,  # better contrast on white
        node_size=node_sizes,
        edgecolors="black",
        linewidths=0.4,
        alpha=0.95
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
    plt.show()


if __name__ == "__main__":
    months_by_residents = read_data("rottedata.xlsx")
    roomie_graph = build_graph(months_by_residents)
    plot_graph(roomie_graph)

