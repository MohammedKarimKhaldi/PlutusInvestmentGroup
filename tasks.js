// Simple in-memory tasks configuration grouped by owner.
// Each task can optionally be linked to a deal via dealId.
const TASKS = [
  {
    id: "t1",
    owner: "Chris",
    dealId: "biolux-seed",
    title: "Prepare Biolux teaser update",
    type: "Document",
    status: "in progress", // in progress | waiting | done
    dueDate: "2026-03-10",
    notes: "Include latest traction metrics from February.",
  },
  {
    id: "t2",
    owner: "Chris",
    dealId: "biolux-seed",
    title: "Schedule partner call with Biolux",
    type: "Meeting",
    status: "waiting",
    dueDate: "2026-03-15",
    notes: "Confirm availability with Biolux CEO.",
  },
  {
    id: "t3",
    owner: "MJ",
    dealId: "iq500-series-a",
    title: "Map top 20 target investors for IQ500",
    type: "Research",
    status: "in progress",
    dueDate: "2026-03-20",
    notes: "",
  },
  {
    id: "t4",
    owner: "MJ",
    dealId: null,
    title: "Update internal fundraising tracker template",
    type: "Internal",
    status: "done",
    dueDate: "2026-03-01",
    notes: "",
  },
];

