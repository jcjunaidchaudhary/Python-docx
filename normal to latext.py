# import re

# def convert_to_latex(expression):
#     # Replace the summation notation with LaTeX \sum
#     expression = re.sub(r'∑_\(([^)]*)\)', r'\\sum_{\1}', expression)
    
#     # Replace the binomial coefficient (n¦k) with LaTeX \binom
#     expression = re.sub(r'\((\w+)¦(\w+)\)', r'\\binom{\1}{\2}', expression)
    
#     # Ensure power notation in LaTeX format with curly braces for superscripts
#     expression = re.sub(r'(\w+)\^(\w+)', r'\1^{\2}', expression)

#     # Return the cleaned LaTeX expression
#     return expression

# # Example input
# input_expression = "(x+a)^n=∑_(k=0)^n(n¦k) x^k a^(n-k)"
# input_expression = "f(x)=a_0+∑_(n=1)^∞ (a_n  cos⁡ nπx/L +b_n  sin⁡ nπx/L ) "
# input_expression = "sin⁡α±sin⁡β=2 sin⁡ 1/2 (α±β)  cos⁡ 1/2 (α∓β)"

# # Convert to LaTeX
# latex_expression = convert_to_latex(input_expression)

# # Output the LaTeX expression
# print("Converted LaTeX Expression: ")
# print(latex_expression)

import re

def convert_to_latex(expression):
    # Replace 'sin' with LaTeX '\sin'
    expression = expression.replace('sin⁡', r'\sin ')
    
    # Replace the ± and ∓ symbols with LaTeX commands
    expression = expression.replace('±', r'\pm ').replace('∓', r'\mp ')
    
    # Replace the fraction 1/2 with LaTeX \frac{1}{2}
    expression = re.sub(r'1/2', r'\\frac{1}{2}', expression)

    # Ensure proper grouping using \left and \right for brackets
    expression = re.sub(r'\((.*?)\)', r'\\left(\1\\right)', expression)
    
    return expression

# Example input
input_expression = "sin⁡α±sin⁡β=2 sin⁡〖1/2 (α±β)〗 cos⁡〖1/2 (α∓β)〗"
input_expression = "f(x)=a_0+∑_(n=1)^∞(a_n  cos⁡〖nπx/L〗+b_n  sin⁡〖nπx/L〗 ) "

# Convert to LaTeX
latex_expression = convert_to_latex(input_expression)

# Output the LaTeX expression
print("Converted LaTeX Expression: ")
print(latex_expression)
