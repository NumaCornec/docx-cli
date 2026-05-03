use std::fmt;

use thiserror::Error;

/// Process exit codes per PRD §10.1.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum ExitCode {
    Success = 0,
    Generic = 1,
    InvalidArgument = 2,
    PreservationImpossible = 3,
    MissingDependency = 4,
    Usage = 64,
}

impl fmt::Display for ExitCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", *self as u8)
    }
}

#[derive(Debug, Error)]
pub enum DocxaiError {
    #[error("error: {0}")]
    Generic(String),

    #[error("error: invalid argument: {0}")]
    InvalidArgument(String),

    #[error("error: preservation impossible: {0}")]
    PreservationImpossible(String),

    #[error("error: missing system dependency: {0}")]
    MissingDependency(String),

    #[error("error: {0} not yet implemented")]
    NotImplemented(&'static str),
}

impl DocxaiError {
    pub fn exit_code(&self) -> ExitCode {
        match self {
            // NotImplemented is treated as generic failure until the verb lands.
            DocxaiError::Generic(_) | DocxaiError::NotImplemented(_) => ExitCode::Generic,
            DocxaiError::InvalidArgument(_) => ExitCode::InvalidArgument,
            DocxaiError::PreservationImpossible(_) => ExitCode::PreservationImpossible,
            DocxaiError::MissingDependency(_) => ExitCode::MissingDependency,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exit_code_repr_matches_prd() {
        assert_eq!(ExitCode::Success as u8, 0);
        assert_eq!(ExitCode::Generic as u8, 1);
        assert_eq!(ExitCode::InvalidArgument as u8, 2);
        assert_eq!(ExitCode::PreservationImpossible as u8, 3);
        assert_eq!(ExitCode::MissingDependency as u8, 4);
        assert_eq!(ExitCode::Usage as u8, 64);
    }

    #[test]
    fn error_kinds_map_to_expected_codes() {
        assert_eq!(
            DocxaiError::Generic("x".into()).exit_code(),
            ExitCode::Generic
        );
        assert_eq!(
            DocxaiError::InvalidArgument("x".into()).exit_code(),
            ExitCode::InvalidArgument
        );
        assert_eq!(
            DocxaiError::PreservationImpossible("x".into()).exit_code(),
            ExitCode::PreservationImpossible
        );
        assert_eq!(
            DocxaiError::MissingDependency("x".into()).exit_code(),
            ExitCode::MissingDependency
        );
        assert_eq!(
            DocxaiError::NotImplemented("snapshot").exit_code(),
            ExitCode::Generic
        );
    }
}
