#include <iostream>
#include <fstream>
#include <xlnt/xlnt.hpp>
#include <string>
#include <vector>

// Function prototypes
void showData(const std::string &filename);
void addHome(const std::string &filename);
void deleteHome(const std::string &filename);
void updateHome(const std::string &filename);
void searchHomes(const std::string &filename);
void ensureExcelFile(const std::string &filename);

void ensureExcelFile(const std::string &filename) {
    std::ifstream file(filename);
    if (!file.good()) {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.active_sheet();
        ws.cell("A1").value("Reference");
        ws.cell("B1").value("Disponibility");
        ws.cell("C1").value("Place");
        ws.cell("D1").value("Sea Distance");
        ws.cell("E1").value("Type");
        wb.save(filename);
        std::cout << "Excel file created with proper columns.\n";
    } else {
        std::cout << "Excel file already exists.\n";
    }
}

int main() {
    std::string filename = "homes.xlsx";
    ensureExcelFile(filename);

    char choice;

    while (true) {
        std::cout << "Choose an option:\n";
        std::cout << "1. Show data\n";
        std::cout << "2. Add home\n";
        std::cout << "3. Delete home\n";
        std::cout << "4. Update home\n";
        std::cout << "5. Search homes\n";
        std::cout << "6. Exit\n";
        std::cin >> choice;

        switch (choice) {
            case '1':
                showData(filename);
                break;
            case '2':
                addHome(filename);
                break;
            case '3':
                deleteHome(filename);
                break;
            case '4':
                updateHome(filename);
                break;
            case '5':
                searchHomes(filename);
                break;
            case '6':
                return 0;
            default:
                std::cout << "Invalid choice. Please try again.\n";
        }
    }
}

void showData(const std::string &filename) {
    xlnt::workbook wb;
    wb.load(filename);
    xlnt::worksheet ws = wb.active_sheet();

    for (auto row : ws.rows(false)) {
        for (auto cell : row) {
            std::cout << cell.to_string() << "\t";
        }
        std::cout << "\n";
    }
}

void addHome(const std::string &filename) {
    xlnt::workbook wb;
    wb.load(filename);
    xlnt::worksheet ws = wb.active_sheet();

    std::string reference, disponibilty, place, sea_distance, type;
    std::cout << "Enter reference: ";
    std::cin >> reference;
    std::cout << "Enter disponibilty (free or rented with date): ";
    std::cin >> disponibilty;
    std::cout << "Enter place: ";
    std::cin >> place;
    std::cout << "Enter sea distance: ";
    std::cin >> sea_distance;
    std::cout << "Enter type: ";
    std::cin >> type;

    int row_num = ws.highest_row() + 1;
    ws.cell("A" + std::to_string(row_num)).value(reference);
    ws.cell("B" + std::to_string(row_num)).value(disponibilty);
    ws.cell("C" + std::to_string(row_num)).value(place);
    ws.cell("D" + std::to_string(row_num)).value(sea_distance);
    ws.cell("E" + std::to_string(row_num)).value(type);

    wb.save(filename);
}

void deleteHome(const std::string &filename) {
    xlnt::workbook wb;
    wb.load(filename);
    xlnt::worksheet ws = wb.active_sheet();

    std::string reference;
    std::cout << "Enter reference to delete: ";
    std::cin >> reference;

    int row_to_delete = -1;
    for (auto row : ws.rows(false)) {
        if (row[0].to_string() == reference) {
            row_to_delete = row[0].reference().row();
            break;
        }
    }

    if (row_to_delete != -1) {
        ws.delete_rows(row_to_delete, 1);
        wb.save(filename);
        std::cout << "Home deleted successfully.\n";
    } else {
        std::cout << "Reference not found.\n";
    }
}

void updateHome(const std::string &filename) {
    xlnt::workbook wb;
    wb.load(filename);
    xlnt::worksheet ws = wb.active_sheet();

    std::string reference, new_disponibilty;
    std::cout << "Enter reference to update: ";
    std::cin >> reference;

    for (auto row : ws.rows(false)) {
        if (row[0].to_string() == reference) {
            char homeFree;
            std::cout << "Is the home free? (Y/N): ";
            std::cin >> homeFree;

            if (homeFree == 'Y' || homeFree == 'y') {
                new_disponibilty = "free";
            } else {
                std::string rent_period;
                std::cout << "Enter rent period (e.g., 20jul-28jul): ";
                std::cin >> rent_period;
                new_disponibilty = rent_period;
            }

            row[1].value(new_disponibilty);
            break;
        }
    }
    wb.save(filename);
}

void searchHomes(const std::string &filename) {
    xlnt::workbook wb;
    wb.load(filename);
    xlnt::worksheet ws = wb.active_sheet();

    std::string place, disponibilty, sea_distance, type;
    std::cout << "Enter place (leave empty if not filtering by place): ";
    std::cin.ignore();
    std::getline(std::cin, place);
    std::cout << "Enter disponibilty (leave empty if not filtering by disponibilty): ";
    std::getline(std::cin, disponibilty);
    std::cout << "Enter sea distance (leave empty if not filtering by sea distance): ";
    std::getline(std::cin, sea_distance);
    std::cout << "Enter type (leave empty if not filtering by type): ";
    std::getline(std::cin, type);

    std::vector<std::vector<xlnt::cell>> matching_rows;

    for (auto row : ws.rows(false)) {
        bool match = true;
        std::vector<xlnt::cell> cells;
        for (auto cell : row) {
            cells.push_back(cell);
        }

        if (!place.empty() && cells[2].to_string() != place) match = false;
        if (!disponibilty.empty() && cells[1].to_string() != disponibilty) match = false;
        if (!sea_distance.empty() && cells[3].to_string() != sea_distance) match = false;
        if (!type.empty() && cells[4].to_string() != type) match = false;

        if (match) {
            matching_rows.push_back(cells); // Store cells in vector
        }
    }

    // Print the search results
    if (!matching_rows.empty()) {
        for (const auto &row : matching_rows) {
            std::cout << "Reference: " << row[0].to_string() 
                      << ", Place: " << row[2].to_string() 
                      << ", Disponibilty: " << row[1].to_string()
                      << ", Sea Distance: " << row[3].to_string() 
                      << ", Type: " << row[4].to_string() << "\n";
        }
    } else {
        std::cout << "No homes matched the search criteria.\n";
    }
}

