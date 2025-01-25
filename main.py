from openai import OpenAI
import pandas as pd
import os



client = OpenAI(
    api_key="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"
)




def gpt_call(prompt):
    """
    Makes a GPT call with the provided prompt.
    """
    try:
        response = client.chat.completions.create(
        messages=[
            {
                "role": "user",
                "content": prompt,
            }
        ],
        model="gpt-4o",
)
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error: {e}")
        return None

def process_excel(file_path):
    """
    Reads an Excel file, processes each row with GPT, and writes the output back.
    """
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Create a new DataFrame to store results
    results = []

    for index, row in df.iterrows():
        input_text = row["Input"]
        prompt = row["Prompt"]

        # Create the final GPT prompt
        # final_prompt = f"{prompt}\nInput: {input_text}"

        # Call GPT
        print(f"Processing row {index + 1}...")
        output = gpt_call(prompt)

        # Append the result to the results list
        results.append({"Input": input_text, "Output": output})

    # Convert the results to a DataFrame
    result_df = pd.DataFrame(results)

    # Write the new DataFrame to the same Excel file
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        result_df.to_excel(writer, sheet_name="GPT_Output", index=False)

    print("Processing complete. Results saved to 'GPT_Output' sheet.")

if __name__ == "__main__":
    # Specify the path to your Excel file
    excel_file_path = "input.xlsx"

    # Process the Excel file
    process_excel(excel_file_path)
