fn main() {
    println!("{:?}", folder_generator::format_project_name("a 123"));
    println!("{:?}", folder_generator::format_project_name("a 123_"));
    println!("{:?}", folder_generator::format_project_name("2025_0004_003_"));
}
