use excel_com::ComApartment;

fn main() {
    let apartment = ComApartment::sta().unwrap();
    std::thread::spawn(move || drop(apartment));
}
