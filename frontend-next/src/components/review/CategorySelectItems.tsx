"use client";

import { Fragment } from "react";
import {
  SelectGroup,
  SelectItem,
  SelectLabel,
  SelectSeparator,
} from "@/components/ui/select";
import { CategoryGroup } from "@/lib/api/types";

type CategorySelectItemsProps = {
  categories: string[];
  categoryGroups?: CategoryGroup[];
};

export function CategorySelectItems({ categories, categoryGroups }: CategorySelectItemsProps) {
  const groups =
    categoryGroups && categoryGroups.length > 0
      ? categoryGroups
      : [{ id: "all", label: "Categories", categories }];

  return (
    <>
      {groups.map((group, groupIndex) => (
        <Fragment key={group.id}>
          {groupIndex > 0 ? <SelectSeparator /> : null}
          <SelectGroup>
            <SelectLabel>{group.label}</SelectLabel>
            {group.categories.map((category) => (
              <SelectItem key={category} value={category}>
                {category}
              </SelectItem>
            ))}
          </SelectGroup>
        </Fragment>
      ))}
    </>
  );
}
