import openpyxl
import networkx as nx
import matplotlib.pyplot as plt
import textwrap
def create_graphs(excel_file):
 wb = openpyxl.load_workbook(excel_file)
 sheet = wb.active
 graphs = []
 current_graph = nx.DiGraph()
 current_tree = None
 for row in sheet.iter_rows(min_row=2):
 tree_id, parent, parent_desc, child, child_desc = [cell.value for cell in row]
 if tree_id and current_tree != tree_id:
 if current_tree is not None:
 graphs.append((current_tree, current_graph))
 current_graph = nx.DiGraph()
 current_tree = tree_id
 current_graph.add_node(parent, description=parent_desc)
 current_graph.add_node(child, description=child_desc)
 current_graph.add_edge(parent, child)
 if current_tree is not None:
 graphs.append((current_tree, current_graph))
 return graphs
def tree_layout(graph, root, width=0.5, vert_gap=0.2, vert_loc=0, xcenter=0.5):
 pos = {root: (xcenter, vert_loc)}
 if graph.out_degree(root) != 0:
 dx = width / 2
 nextx = xcenter - (dx * (graph.out_degree(root) - 1)) / 2
 for child in graph.neighbors(root):
 pos.update(tree_layout(graph, child, width=dx, vert_gap=vert_gap, 
vert_loc=vert_loc-vert_gap, xcenter=nextx))
 nextx += dx
 return pos
def visualize_graphs(graphs):
 rows = (len(graphs) + 1) // 2
 cols = 2 # Display 2 trees per page
 for tree_number, ((tree_id, graph), page_start) in enumerate(zip(graphs, range(0, 
len(graphs), cols)), start=1):
 page_graphs = graphs[page_start:page_start + cols]
 num_graphs = len(page_graphs)
 rows = 1
 fig, axs = plt.subplots(rows, num_graphs, figsize=(15, 6))
 if num_graphs == 1:
 axs = [axs] # To handle a single subplot case
 for i, (ax, (tree_id, graph)) in enumerate(zip(axs, page_graphs)):
 root_nodes = [n for n, d in graph.in_degree() if d == 0]
 for root_node in root_nodes:
 pos = tree_layout(graph, root_node, xcenter=0.5) # Center each root node
 last_parent = None
 node_labels = {}
 labels = {}
 for node, desc in graph.nodes(data='description'):
 x, y = pos[node]
 desc_wrapped = '\n'.join(textwrap.wrap(desc, width=20))
 num_lines = len(desc_wrapped.split('\n'))
 box_height = 0.015 * num_lines
 desc_x = x - (box_height / 2) # Adjust this value to move descriptions 
further to the side
 if graph.out_degree(node) == 0:
 ax.text(x, y - 0.04 - 0.6 * box_height, desc_wrapped, fontsize=8, 
ha='center', va='center', zorder=1)
 else:
 ax.text(desc_x, y, desc_wrapped, fontsize=8, ha='left', va='center', 
zorder=1)
 ax.add_patch(plt.Rectangle((desc_x, y - 0.04 - 0.6 * box_height), 0.16, 
box_height, color='white', edgecolor='gray', linewidth=1, zorder=0))
 if last_parent and node != last_parent:
 node_labels[last_parent] = "\n".join(labels.values())
 labels = {}
 labels[node] = node # Store node name in labels
 last_parent = node
 if labels:
 node_labels[last_parent] = "\n".join(labels.values())
 nx.draw(graph, pos, with_labels=False, node_size=1500, 
node_color='lightblue', font_size=10, font_color='black', ax=ax)
 nx.draw_networkx_labels(graph, pos, labels=node_labels, font_size=8, 
verticalalignment='center', ax=ax)
 ax.set_title(f"Tree {tree_id}")
 ax.axis('off')
 plt.tight_layout()
 plt.show()
if __name__ == '__main__':
 excel_file = 'excelsheet.xlsx' # Provide the correct path to your Excel sheet
 graphs = create_graphs(excel_file)
 visualize_graphs(graphs)
