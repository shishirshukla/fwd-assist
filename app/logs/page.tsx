import { LogsList } from "@/components/logs-list";

export default function LogsPage() {
  return (
    <div className="h-full overflow-y-auto bg-background">
      <LogsList />
    </div>
  );
}
