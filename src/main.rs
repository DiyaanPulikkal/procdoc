use clap::Parser;
use proc_doc::Instruction;
use proc_doc::verify_args;

fn main(){
    // Parse Arguments
    let instruction:Instruction = Instruction::parse();

    // Verify if the arguments are valid for the program
    if instruction.input_path.is_none(){       
        eprintln!("Please enter the path of input file.");
        std::process::exit(0);
    }

    verify_args(instruction);

    println!("Success!")

}




