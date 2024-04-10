def get_placeholder_values(xy_count, naming_template):
    placeholder_values = {}
    for i in range(xy_count):
        xy_name = f"XY{i+1:02}"
        placeholder_values[xy_name] = {}
        for placeholder in range(1, naming_template.count("{") - naming_template.count("{C}") + 1):
            value = input(f"Enter placeholder {{{placeholder}}} for {xy_name}: ")
            placeholder_values[xy_name][f'{placeholder}'] = value
    return placeholder_values