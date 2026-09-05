# Universal agent repository contract

This repository is operated by NODO, Ninja, and registered Rodolfo OS agents.
Agents may inspect, configure, implement, document, test, branch, review, hand
off, and open PRs autonomously within an accepted task envelope.

Before work, read the SOR, README, runbooks, architecture/ADR, tests, runtime
records, and deployment documentation. Search GitHub, VM checkouts, local
workspaces, connectors, and existing artifacts first. Reuse the canonical
repository, checkout, dependency cache, workflow, connector, and ledger.
Create a private traceability repository only when no suitable repository
exists; check collisions and record provenance and the runtime-authority
relationship. Never copy secrets, sessions, logs, runtime state, customer or
personal data into GitHub.

Use task-nature, least-privilege credentials through approved secure injection.
Profiles must be identity-verified, environment-scoped, revocable or
time-bounded, and auditable. If access or repository creation is unavailable,
return `PENDING_CREDENTIAL_PROFILE` or `PENDING_REPOSITORY_CONNECTOR`; never
invent, print, commit, or transmit credentials.

Every material task declares complexity and quality risks; intake and Task
Model; T0/T1/U1/T2-T3 routing; deterministic opportunities and share;
reused/promoted artifacts; provenance and uncertainty; escalation trigger,
maximum tier and return condition; expected cost/quota; rollback; and final
verification against the original request. Report tier by stage, reuse,
escalation reason, return to normal tier, and final synthesis. Test both
underpowered-model failure and unnecessary escalation.

Agents may communicate with registered agents, request tasks, assign bounded
peer work, review diffs, and hand off work while preserving task_id,
parent_task_id, provenance, acceptance evidence, and return condition. Avoid
duplicate resources and preserve dirty work.

Stop at `DEPLOYMENT_READY`. Deployments, restarts, production mutations,
public releases, IAM changes, secret changes, destructive actions, and
irreversible external actions require Rodolfo's explicit authorization naming
the exact task, repository, commit or PR, target environment, and action.
Green tests, merge, a workflow, or a standing migration profile never replace
that authorization.
