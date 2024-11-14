import { useState } from "react";
import ReadExcel from "./ReadExcel";
import "./App.css";

function App() {
  const [count, setCount] = useState(0);

  return (
    <>
      <ReadExcel />
    </>
  );
}

export default App;
