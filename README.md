# OTornotOT — Work Task Planner

We think everyone has the needs to organise their working tasks efficiently.
The hardest is the creating the planner. Yet, people keep struggle with how to keep their tasks well organised.
We know that a planner is a burden if people always get stress in their works, they are not good at planning, and there are too many things to do to plan.
Hence, we've created this planner in a way of easy and fun with a hope people can have some laughing and enjoying moments with planning.

1. **Daily tasks** — Top 3 should do, allocate tasks to morning & afternoon, Reschedule when you are too overload.
2. **Progress** — See how long period tasks may affect your week tasks and see how many stuffs are piled each day.
3. **Emails** — Get tasks auto from linked emails, receive alerts of urgent emails, sync your calendar.
4. **Tutorials** — Have fun when get onboarding.

## Live Demo

[https://endpoint-9f19ce47-3400-402b-8e97-05d13f2bb9b0.agentbase-runtime.aiplatform.vngcloud.vn](https://endpoint-9f19ce47-3400-402b-8e97-05d13f2bb9b0.agentbase-runtime.aiplatform.vngcloud.vn)

## Tech Stack

- **Frontend**: React (single HTML file)
- **Backend**: Python + Starlette + Uvicorn
- **Deployment**: GreenNode AgentBase (runtime size 2x4)
- **Container**: Docker (linux/amd64)

## Project Structure

```
OTornotOT/
├── main.py              # Starlette app — serves TaskPlanner.html
├── TaskPlanner.html     # React frontend (single-file app)
├── requirements.txt     # Python dependencies
├── Dockerfile           # Container definition
└── .env.example         # Environment variable template
```

## Environment Variables

Copy `.env.example` to `.env` and fill in your values:

```
GREENNODE_CLIENT_ID=
GREENNODE_CLIENT_SECRET=
LLM_API_KEY=
LLM_BASE_URL=https://maas-llm-aiplatform-hcm.api.vngcloud.vn/v1
LLM_MODEL=qwen/qwen3-5-27b
```

## Run Locally

```bash
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8080
```

Open [http://localhost:8080](http://localhost:8080)

## Deploy to GreenNode AgentBase

```bash
docker build --platform linux/amd64 -t vcr.vngcloud.vn/<project>/<repo>:latest .
docker push vcr.vngcloud.vn/<project>/<repo>:latest
```

Then update the runtime via the AgentBase console or API.
