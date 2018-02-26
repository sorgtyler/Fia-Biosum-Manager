# fia-biosum-manager
User interface and main code repository for Biosum.

This branch is for developing functionality to create MS Access databases as inputs for FVS.

The following files were modified for creating FVS input databases:
* fvs\_input.cs
  - New functions:
    - CopyFVSBlankDatabaseToFVSInDir()
    - CreateTablesLinksToFVSIn()
    - CreateFVSInputDbLOC()
    - CreateFVSStandInit()
    - GenerateSiteIndexAndSiteSpeciesSQL(): The sequential code from the CreateSLF method for calculating SiteIndex/SiteSpecies
    - CreateFVSTreeInit()
  - Modified functions:
    - Start(): Calls a Start helper function StartFVSInAccdb instead
* Tables.cs FVS nested class
  - New Functions:
    - CreateFVSInputStandInitTable()
    - CreateFVSInputStandInitTableIndexes()
    - CreateFVSInputStandInitTableSQL()
    - CreateFVSInputTreeInitWorkTable()
    - CreateFVSInputTreeInitWorkTableIndexes()
    - CreateFVSInputTreeInitWorkTableSQL()
* Queries.cs
  - New nested classes:
    - FVS.FVSInput.StandInit
      - Many query strings are defined here for importing stand data and manipulating it.
    - FVS.FVSInput.TreeInit
      - Ditto for tree data.

As a consequence of this new approach, the `fvs_tree_id` column used to track trees throughout the life of a BioSum project was redefined. Accordingly, the following parts of BioSum related to fvs_tree_id need to be updated:
* uc\_plot\_input.cs should define `fvs_tree_id` as `Master.Tree.Subp*1000 + Master.Tree.Tree`.
* uc\_fvs\_output.cs should join the FVS Tree and BioSum Tree records on `Stand_ID = Biosum_Cond_Id AND Tree_ID = fvs_tree_id` because `fvs_tree_id` is only unique per stand.
* Queries.cs should:
  - use the `fvs_tree_id` column when populating the `FVSIn.FVS_TreeInit` table.
  - create a compound primary key of `(biosum_cond_id, fvs_tree_id)` in the FVSOut POST-Append audit `tree_fvs_tree_id_work_table`.
  - update the join conditions described in the uc\_fvs\_output.cs point.
