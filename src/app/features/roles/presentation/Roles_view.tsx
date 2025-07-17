interface RolesViewProps {
  dynamicMode: boolean;
}

export default function RolesView({ dynamicMode }: RolesViewProps) {
  return (
    <div>
      <h1>Roles View</h1>
      <p>Dynamic Mode: {dynamicMode ? "Enabled" : "Disabled"}</p>
      {/* Additional content can be added here based on the dynamic mode */}
    </div>
  );
}
