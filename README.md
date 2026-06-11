all_data = []
errors = []

for i, file in enumerate(files, start=1):
    try:
        file_type = classify_file(file)

        if file_type == "TYPE_1_HEADER":
            df = process_type_1(file)
        else:
            df = process_type_2(file)

        all_data.append(df)

        print(f"{i}/{len(files)} processed - {file_type} - {file.name}")

    except Exception as e:
        errors.append((file.name, str(e)))
        print("ERROR:", file.name, "->", e)

print("Processed files:", len(all_data))
print("Errors:", len(errors))
