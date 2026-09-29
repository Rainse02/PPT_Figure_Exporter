use std::env;
use std::fs;
use std::process::ExitCode;

use figure_export_core::{process_request, CoreError, Request, Response};

fn main() -> ExitCode {
    match run() {
        Ok(response) => {
            match serde_json::to_string(&response) {
                Ok(json) => println!("{json}"),
                Err(error) => {
                    eprintln!("{error}");
                    return ExitCode::from(1);
                }
            }
            ExitCode::SUCCESS
        }
        Err(error) => {
            let response = error_response(&error);
            match serde_json::to_string(&response) {
                Ok(json) => println!("{json}"),
                Err(serialization_error) => eprintln!("{serialization_error}"),
            }
            eprintln!("{error}");
            ExitCode::from(1)
        }
    }
}

fn run() -> Result<Response, CoreError> {
    let mut arguments = env::args_os().skip(1);
    let request_flag = arguments.next().ok_or_else(|| {
        CoreError::InvalidRequest("用法：figure-export-core.exe --request <JSON路径>".to_owned())
    })?;
    if request_flag != "--request" {
        return Err(CoreError::InvalidRequest(
            "用法：figure-export-core.exe --request <JSON路径>".to_owned(),
        ));
    }
    let request_path = arguments
        .next()
        .ok_or_else(|| CoreError::InvalidRequest("缺少 JSON 请求路径".to_owned()))?;
    if arguments.next().is_some() {
        return Err(CoreError::InvalidRequest(
            "只允许提供一个 --request JSON 路径".to_owned(),
        ));
    }
    let data = fs::read(&request_path)
        .map_err(|error| CoreError::Input(format!("无法读取请求 JSON：{error}")))?;
    let request = serde_json::from_slice::<Request>(&data)
        .map_err(|error| CoreError::InvalidRequest(format!("请求 JSON 无效：{error}")))?;
    process_request(&request)
}

fn error_response(error: &CoreError) -> serde_json::Value {
    serde_json::json!({
        "ok": false,
        "error": error.to_string(),
    })
}
