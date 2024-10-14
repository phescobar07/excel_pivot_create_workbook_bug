import * as React from "react";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC = () => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      hello world
      {/* <er logo="assets/logo-filled.png" title={props.title} message="Welcome" /> */}
      {/* <HeroList message="Discover what this add-in can do for you today!" items={listItems} /> */}
      {/* <TextInsertion insertText={insertText} /> */}
    </div>
  );
};

export default App;
