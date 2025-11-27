fn main() {
    // Only run on Windows
    #[cfg(windows)]
    {
        let mut res = winres::WindowsResource::new();
        res.set_icon("icon.ico");
        res.compile().expect("Failed to compile Windows resources");
    }
}
