// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const searchInTree = (tree : any, value : unknown, key : string, subKey?: string) => {
  let treeValue = null;
  if (subKey) {
    treeValue = tree[key][subKey];
  } else {
    treeValue = tree[key];
  }

  if (treeValue === value) {
    return tree;
  }

  if (tree._children)
    tree.children = tree._children;

  for (const child of tree.children) {
    const found = searchInTree(child, value, key, subKey);

    if (found) {
      if (found._children)
        found.children = found._children;

      return found;
    }
  }
};
