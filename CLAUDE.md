This project contains the codebase for the data collection tools
associated with a series of Delphi-style workshops to optimize public
health programs in Mozambique. They also contain supporting generated
static content describing the programs to be optimized.

As a general rule, no true secrets or personal information should be
committed to the git repository. Secrets needed for the dashboard are
generated temporarily and stored on the streamlit.io server. Secrets
needed for deployment of the Kobo questionnaire are stored locally in
config.env.

The project is organized into worktrees for each workshop:

INS_delphi - malaria - 1st workshop
|
|- INS_delphi_f2 - HIV/TB/SMI - 2nd round of workshops
|   |- INS_delphi_hiv - HIV - 1st workshop (SMI and TB were postponed)
|   \- INS_delphi_tb - TB - 1st workshop
\-- catalog-pages - branch for supporting Git Pages deployment of static content

There is no requirement that all changes make it back to the root
branch, but changes to older branches should generally be merged
forward to support future workshops.

