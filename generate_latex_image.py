import matplotlib.pyplot as plt
from matplotlib import rc
import os

def generate_latex_image(latex_expression, image_filename):
    # Configure LaTeX rendering
    rc('text', usetex=True)
    rc('font', family='serif')

    # Create a figure with dynamic size
    fig = plt.figure()

    # Hide the axes
    ax = fig.add_subplot(111)
    ax.axis('off')

    # Add LaTeX text to the center of the plot
    rendered_text = ax.text(0.5, 0.5, f"${latex_expression}$", size=20, ha='center', va='center')

    # Adjust the figure size based on the content width
    fig.canvas.draw()
    bbox = rendered_text.get_window_extent(renderer=fig.canvas.get_renderer())
    width, height = bbox.width / fig.dpi, bbox.height / fig.dpi

    # Set the figure size based on text size
    fig.set_size_inches(width, height)

    # Save the image with no padding and tight bounding box
    plt.savefig(image_filename, bbox_inches='tight', pad_inches=0)

    # Close the figure to free memory
    plt.close(fig)

# Example usage:
latex_expression = r"\int_a^b x^2 \, dx"
image_filename = "latex_expression_dynamic_size.png"

generate_latex_image(latex_expression, image_filename)
