import { TaskPane } from "@/components/task-pane";

export default function TaskPanePage() {
  return (
    <div className="min-h-screen bg-background p-4">
      <h1 className="mb-1 text-base font-semibold">Forward Guard</h1>
      <p className="mb-4 text-xs text-muted-foreground">
        Priority, End Date, and Category
      </p>
      <TaskPane />
    </div>
  );
}
