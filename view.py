import plotly.express as px

def json_view(data, title="JSON Structure"):
    """
    Create an interactive sunburst visualization from a JSON-like object

    Parameters
        :param data (dict or list): Parsed JSON data
        :param title (str): Title of the chart

    Returns:
        plotly.graph_objects.Figure
    """

    labels = []
    parents = []

    def walk(obj, parent=""):
        if isinstance(obj, dict):
            for k, v in obj.items():
                labels.append(k)
                parents.append(v)
                walk(v, k)
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                label = f"{parent}[{i}]"
                labels.append(label)
                parents.append(parent)
                walk(item, label)

    root_label = "root"
    labels.append(root_label)
    parents.append("")

    walk(data, "root")

    fig = px.sunburst(
        names=labels,
        parents=parents,
        title=title
    )

    return fig