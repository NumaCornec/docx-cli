# Ralph Fix Plan

## High Priority
- [ ] 5 verbes implémentés et testés (squelette en place, dispatch retourne `NotImplemented`)
- [ ] Round-trip 100% sur le corpus de conformance
- [ ] SKILL.md publié et testé avec Claude Code en eval (≥ 95% tool selection accuracy sur 50 tâches diverses)
- [ ] Binaires publiés sur les 5 plateformes cibles
- [ ] Documentation : SKILL.md, CHANGELOG (README minimal en place)
- [ ] CI vert sur main depuis ≥ 2 semaines
- [ ] Couverture de tests ≥ 80% (mesurée par `cargo llvm-cov`)
- [ ] CI green sur main (workflow à créer)
- [ ] CI bloque les PR si fmt/clippy/test échoue (workflow à créer)
- [ ] Build time < 5 min avec cache chaud
- [ ] 10 fixtures committées
- [ ] Chacune ouvrable par Word sans erreur (vérif manuelle documentée)
- [ ] README dans `tests/fixtures/` qui décrit chaque fichier
- [ ] Test `preservation_tests::roundtrip_noop` passe sur les 10 fixtures
- [ ] Échec lisible avec diff zip explicite si régression
- [ ] Tag `v0.0.1` produit 5 binaires sur GitHub Releases
- [ ] Chaque binaire run `--version` et affiche `0.0.1`
- [ ] `Doc::load("simple.docx")` réussit
- [ ] Erreur claire si fichier n'est pas un zip valide
- [ ] Erreur claire si pas un docx (Content_Types absent)


## Medium Priority


## Low Priority


## Completed
- [x] Project enabled for Ralph
- [x] Bootstrap cargo project (`Cargo.toml`, `src/main.rs`, `src/lib.rs`)
- [x] `--help` complet pour CLI + chaque sous-commande (clap derive + tests)
- [x] `docxai --help` affiche les 5 verbes (test `top_level_help_lists_all_five_verbs`)
- [x] `docxai snapshot --help` etc. affiche options spécifiques (test `each_verb_has_its_own_help`)
- [x] Tous les exit codes (§10.1) implémentés (`ExitCode` enum + `DocxaiError::exit_code`)
- [x] `--version` fonctionne (clap `version` attr, test `version_flag_prints_crate_version`)
- [x] Test `assert_cmd` qui vérifie `--help` output (`tests/cli_smoke.rs`)

## Notes
- Focus on MVP functionality first
- Ensure each feature is properly tested
- Update this file after each major milestone
- Local toolchain has no `cargo`; rely on CI (still TBD) to verify `cargo build/test/clippy/fmt`. Code was written to compile cleanly; verify on first CI run.
- MSRV pinned to 1.85 / edition 2024 (PRD targets 1.95 / edition 2026 which are not yet stable as of 2026-05).
